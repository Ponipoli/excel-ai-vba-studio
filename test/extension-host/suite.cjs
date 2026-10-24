const assert = require('node:assert/strict');
const { readFile, writeFile } = require('node:fs/promises');
const { join } = require('node:path');
const vscode = require('vscode');
const JSZip = require('jszip');
const { chromium } = require('playwright-core');

async function until(check, label, timeout = 60_000) {
  const deadline = Date.now() + timeout;
  let lastError;
  while (Date.now() < deadline) {
    try {
      const result = await check();
      if (result) return result;
    } catch (error) { lastError = error; }
    await new Promise(resolve => setTimeout(resolve, 150));
  }
  throw new Error(`Timed out: ${label}${lastError ? ` (${lastError.message})` : ''}`);
}

async function spreadsheetFrame(browser) {
  return until(async () => {
    for (const context of browser.contexts()) {
      for (const page of context.pages()) {
        for (const frame of page.frames()) {
          if (await frame.locator('.excel-viewer .x-spreadsheet-formula-bar-name').isVisible()) {
            return frame;
          }
        }
      }
    }
    return null;
  }, 'spreadsheet webview and formula bar');
}

async function selectCell(frame, address) {
  const cellAddress = frame.getByRole('textbox', { name: 'Cell address', exact: true });
  await cellAddress.fill(address);
  await cellAddress.press('Enter');
}

async function setCell(frame, address, value) {
  await selectCell(frame, address);
  const content = frame.getByRole('textbox', { name: 'Cell content', exact: true });
  await content.fill(value);
  await content.press('Enter');
}

async function openWorkbook(browser, filePath) {
  await vscode.commands.executeCommand('vscode.openWith', vscode.Uri.file(filePath), 'excelAiVbaStudio.officeViewer');
  return spreadsheetFrame(browser);
}

async function closeEditors() {
  await vscode.commands.executeCommand('workbench.action.closeAllEditors');
}

exports.run = async function run() {
  const { ExcelJS } = require(process.env.EXCEL_SMOKE_EXCELJS_FIXTURE);
  const workspace = process.env.EXCEL_SMOKE_WORKSPACE;
  assert.ok(workspace, 'synthetic workspace must be provided');
  const actualWorkspace = vscode.workspace.workspaceFolders[0].uri.fsPath;
  // VS Code normalizes Windows drive-letter casing in URI.fsPath.
  assert.equal(
    process.platform === 'win32' ? actualWorkspace.toLowerCase() : actualWorkspace,
    process.platform === 'win32' ? workspace.toLowerCase() : workspace,
  );
  const extension = vscode.extensions.getExtension('steph-tools.excel-ai-vba-studio');
  assert.ok(extension, 'the VSIX must be installed in the isolated profile');
  assert.ok(!extension.extensionPath.includes(join('test', 'extension-host')), 'test the installed product');
  await extension.activate();
  assert.equal(extension.isActive, true);
  const commands = await vscode.commands.getCommands(true);
  assert.ok(commands.includes('excelAiVbaStudio.newWorkbook'), 'new-workbook command must be registered');

  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('Synthetic');
  sheet.getCell('A1').value = 'original';
  sheet.getCell('B1').value = 2;
  const path = join(workspace, 'portable-smoke.xlsx');
  await writeFile(path, Buffer.from(await workbook.xlsx.writeBuffer()));
  const original = await readFile(path);
  const browser = await chromium.connectOverCDP(process.env.EXCEL_SMOKE_CDP_ENDPOINT);
  try {
    let frame = await openWorkbook(browser, path);
    assert.equal(await frame.locator('.excel-load-error').count(), 0);
    await selectCell(frame, 'A1');
    await until(async () => await frame.getByRole('textbox', { name: 'Cell content', exact: true }).inputValue() === 'original', 'initial cell content');
    await setCell(frame, 'A1', 'portable-edited');
    await setCell(frame, 'B1', '=SUM(2,3)');
    await until(() => vscode.window.tabGroups.activeTabGroup.activeTab?.isDirty, 'custom editor dirty event');
    await vscode.commands.executeCommand('workbench.action.files.save');
    await until(async () => {
      if (vscode.window.tabGroups.activeTabGroup.activeTab?.isDirty) return false;
      return !(await readFile(path)).equals(original);
    }, 'custom editor save');
    const saved = new ExcelJS.Workbook();
    await saved.xlsx.load(await readFile(path));
    assert.equal(saved.getWorksheet('Synthetic').getCell('A1').value, 'portable-edited');
    assert.equal(saved.getWorksheet('Synthetic').getCell('B1').formula, 'SUM(2,3)');
    await closeEditors();
    frame = await openWorkbook(browser, path);
    await selectCell(frame, 'A1');
    assert.equal(await frame.getByRole('textbox', { name: 'Cell content', exact: true }).inputValue(), 'portable-edited');
    const screenshot = join(process.env.EXCEL_SMOKE_SCREENSHOTS, `${process.platform}-${process.arch}-vscode-${vscode.version}.png`);
    await frame.locator('.excel-viewer').screenshot({ path: screenshot });
    console.log(`Synthetic workbook screenshot: ${screenshot}`);
    await closeEditors();

    const createdPath = join(workspace, 'created-through-command.xlsx');
    const creation = vscode.commands.executeCommand('excelAiVbaStudio.newWorkbook');
    const filename = await until(async () => {
      for (const context of browser.contexts()) {
        for (const page of context.pages()) {
          const input = page.locator('.quick-input-widget .quick-input-box input');
          if (await input.isVisible()) return input;
        }
      }
      return null;
    }, 'new workbook save dialog');
    await filename.fill(createdPath);
    await filename.press('Enter');
    await creation;
    frame = await spreadsheetFrame(browser);
    await setCell(frame, 'A1', 'created-in-vscode');
    await until(() => vscode.window.tabGroups.activeTabGroup.activeTab?.isDirty, 'new workbook dirty event');
    await vscode.commands.executeCommand('workbench.action.files.save');
    await until(() => !vscode.window.tabGroups.activeTabGroup.activeTab?.isDirty, 'new workbook save');
    const created = new ExcelJS.Workbook();
    await created.xlsx.load(await readFile(createdPath));
    assert.equal(created.getWorksheet('Sheet1').getCell('A1').value, 'created-in-vscode');
    await closeEditors();

    if (process.platform !== 'win32') {
      for (const tool of extension.packageJSON.contributes.languageModelTools) {
        assert.equal(tool.when, 'excelAiVbaStudio.nativeExcel', 'native tool UI must require the Windows capability');
      }
      // Newer VS Code exposes disabled contributions in lm.tools as well.
      // Verify runtime unavailability instead of mistaking metadata for an implementation.
      const nativeReadTool = vscode.lm?.tools.find(tool => tool.name === 'excel_ai_vba_readWorkbook');
      if (nativeReadTool) {
        await assert.rejects(
          vscode.lm.invokeTool(nativeReadTool.name, {
            input: { workbookPath: path, includeVba: false, format: 'json' },
            toolInvocationToken: undefined,
          }),
          /not.*implementation.*registered/i,
          'native tool invocation must fail before any workbook automation starts',
        );
      }
      const archive = await JSZip.loadAsync(original);
      const types = await archive.file('[Content_Types].xml').async('string');
      archive.file('[Content_Types].xml', types.replace(
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml',
        'application/vnd.ms-excel.sheet.macroEnabled.main+xml',
      ));
      const macroPath = join(workspace, 'read-only.xlsm');
      await writeFile(macroPath, await archive.generateAsync({ type: 'nodebuffer' }));
      const before = await readFile(macroPath);
      frame = await openWorkbook(browser, macroPath);
      assert.equal(await frame.getByRole('textbox', { name: 'Cell content', exact: true }).isEditable(), false);
      assert.ok(await frame.locator('.excel-readonly-banner').isVisible(), 'macro workbook must explain read-only mode');
      await frame.locator('.excel-viewer').screenshot({
        path: join(process.env.EXCEL_SMOKE_SCREENSHOTS, `${process.platform}-${process.arch}-vscode-${vscode.version}-read-only.png`),
      });
      await closeEditors();
      assert.deepEqual(await readFile(macroPath), before, 'read-only workbook bytes must be unchanged');
    }
    console.log('Installed VSIX smoke passed: activation, custom editor rendering, cell/formula edit, host save, disk verification, reopen, new-workbook command, and portable macro protection.');
  } finally {
    await closeEditors();
    await browser.close();
  }
};
