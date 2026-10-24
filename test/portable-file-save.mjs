import assert from 'node:assert/strict';
import { createHash } from 'node:crypto';
import { build } from 'esbuild';
import { mkdtemp, readFile, writeFile, chmod, stat, symlink, link, readdir, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';
import { pathToFileURL } from 'node:url';

const root = resolve(import.meta.dirname, '..');
const temp = await mkdtemp(join(tmpdir(), 'excel-portable-save-'));
const digest = value => createHash('sha256').update(value).digest('hex');
try {
  const compiled = join(temp, 'save.mjs');
  await build({ entryPoints: [join(root, 'src/provider/portableFileSave.ts')], outfile: compiled,
    bundle: true, format: 'esm', platform: 'node', logLevel: 'silent' });
  const { createPortableFile, savePortableFile } = await import(pathToFileURL(compiled).href);
  const file = join(temp, 'synthetic.xlsx');
  const original = Buffer.from('synthetic original bytes');
  const edited = Buffer.from('synthetic edited bytes');
  assert.equal(await createPortableFile(file, original), digest(original));
  await assert.rejects(createPortableFile(file, edited), /already exists/);
  assert.deepEqual(await readFile(file), original);
  // Staging logic is POSIX-specific; Windows keeps the NTFS-preserving writer.
  if (process.platform !== 'win32') {
    await chmod(file, 0o640);
    assert.equal(await savePortableFile(file, edited, digest(original)), digest(edited));
    assert.deepEqual(await readFile(file), edited);
    assert.equal((await stat(file)).mode & 0o777, 0o640);
    const backups = join(temp, '.excel-ai-vba-backups');
    const recovery = await readdir(backups);
    assert.equal(recovery.length, 1);
    assert.deepEqual(await readFile(join(backups, recovery[0])), original);
    await assert.rejects(savePortableFile(file, original, digest(original)), /changed outside/);
    assert.deepEqual(await readFile(file), edited);
    await assert.rejects(savePortableFile(file, original, ''), /snapshot is missing/);
    await chmod(file, 0o444);
    await assert.rejects(savePortableFile(file, original, digest(edited)), /read-only/);
    await chmod(file, 0o640);
    const alias = join(temp, 'alias.xlsx');
    await symlink(file, alias);
    await assert.rejects(savePortableFile(alias, original, digest(edited)), /symbolic or hard links/);
    const hard = join(temp, 'hard.xlsx');
    await link(file, hard);
    await assert.rejects(savePortableFile(file, original, digest(edited)), /symbolic or hard links/);
    await rm(hard);
    await assert.rejects(createPortableFile(join(backups, 'new.xlsx'), edited), /Recovery backups/);
    const directoryLink = join(temp, 'linked-parent');
    await symlink(temp, directoryLink);
    await assert.rejects(createPortableFile(join(directoryLink, 'new.xlsx'), edited), /symbolic links/);
    // Concurrent edits from two editor sessions cannot both replace the baseline.
    const results = await Promise.allSettled([
      savePortableFile(file, Buffer.from('session one'), digest(edited)),
      savePortableFile(file, Buffer.from('session two'), digest(edited)),
    ]);
    assert.ok(results.some(result => result.status === 'fulfilled'));
    assert.ok(results.some(result => result.status === 'rejected'), 'stale parallel save must be rejected');
  }
  assert.deepEqual((await readdir(temp)).filter(name => name.startsWith('.excel-ai-save-')), []);
  console.log('Portable file save passed: exclusive create, staged replacement, recovery, mode preservation and stale/link/read-only refusals.');
} finally {
  await rm(temp, { recursive: true, force: true });
}
