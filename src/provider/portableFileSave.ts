import { createHash, randomUUID } from 'node:crypto';
import { constants, promises as fs } from 'node:fs';
import type { FileHandle } from 'node:fs/promises';
import { dirname, extname, join, resolve } from 'node:path';

const MAX_PORTABLE_FILE_BYTES = 128 * 1024 * 1024;
const BACKUP_DIRECTORY = '.excel-ai-vba-backups';
const pendingSaves = new Map<string, Promise<string>>();

function sameLocalPath(left: string, right: string): boolean {
    // VS Code normalizes drive letters to lowercase; Windows realpath can return
    // uppercase volume names. This is the same directory, not a symlink redirect.
    const normalizeDrive = (value: string) => value.replace(/^[a-z]:/i, drive => drive.toUpperCase());
    return process.platform === 'win32' ? normalizeDrive(left) === normalizeDrive(right) : left === right;
}

function hash(bytes: Uint8Array): string {
    return createHash('sha256').update(bytes).digest('hex');
}

function validateBytes(bytes: Uint8Array): void {
    if (!(bytes instanceof Uint8Array) || bytes.byteLength > MAX_PORTABLE_FILE_BYTES) {
        throw new Error('Spreadsheet save exceeds the 128 MiB safety limit.');
    }
}

async function localParent(filePath: string): Promise<string> {
    const fullPath = resolve(filePath);
    const parent = dirname(fullPath);
    if (fullPath.split(/[\\/]/).includes(BACKUP_DIRECTORY)) {
        throw new Error('Recovery backups cannot be overwritten by the spreadsheet editor.');
    }
    // A different real parent would let a symlink redirect a staged replacement.
    if (!sameLocalPath(await fs.realpath(parent), parent)) {
        throw new Error('Portable saves require a local path without symbolic links.');
    }
    return parent;
}

async function readStable(handle: FileHandle): Promise<Buffer> {
    const before = await handle.stat();
    if (!before.isFile() || before.size > MAX_PORTABLE_FILE_BYTES) {
        throw new Error('Spreadsheet source is not a regular file within the 128 MiB limit.');
    }
    const bytes = Buffer.alloc(before.size);
    let offset = 0;
    while (offset < bytes.length) {
        const result = await handle.read(bytes, offset, bytes.length - offset, offset);
        if (result.bytesRead === 0) throw new Error('Spreadsheet changed while it was being read.');
        offset += result.bytesRead;
    }
    const extra = await handle.read(Buffer.alloc(1), 0, 1, offset);
    const after = await handle.stat();
    if (extra.bytesRead || before.size !== after.size || before.mtimeMs !== after.mtimeMs
        || before.ctimeMs !== after.ctimeMs) {
        throw new Error('Spreadsheet changed while it was being read.');
    }
    return bytes;
}

async function writeExclusive(filePath: string, bytes: Uint8Array, mode: number): Promise<void> {
    const handle = await fs.open(filePath, 'wx', mode);
    try {
        await handle.writeFile(bytes);
        await handle.sync();
    } finally {
        await handle.close();
    }
}

/** Create without replacing a file selected accidentally in Save As. */
export async function createPortableFile(filePath: string, bytes: Uint8Array): Promise<string> {
    validateBytes(bytes);
    const fullPath = resolve(filePath);
    const parent = await localParent(fullPath);
    const staging = await fs.mkdtemp(join(parent, '.excel-ai-save-'));
    const stagePath = join(staging, 'workbook');
    try {
        await writeExclusive(stagePath, bytes, 0o600);
        if (!sameLocalPath(await fs.realpath(parent), parent)) throw new Error('Save destination changed.');
        // link(2) installs the complete file atomically and fails if the target exists.
        // Unlike rename, this can never clobber a destination created during Save As.
        await fs.link(stagePath, fullPath);
        return hash(bytes);
    } catch (error) {
        if ((error as NodeJS.ErrnoException).code === 'EEXIST') {
            throw new Error('The destination already exists. Choose a new name to preserve that file.');
        }
        throw error;
    } finally {
        await fs.rm(staging, { recursive: true, force: true });
    }
}

/**
 * POSIX staged save tied to the bytes last delivered by the host. Windows uses
 * its existing file writer so NTFS origin streams are not lost by a POSIX rename.
 * The recovery copy preserves the previous bytes even if a later writer races us.
 */
export function savePortableFile(
    filePath: string,
    bytes: Uint8Array,
    expectedSha256: string,
): Promise<string> {
    const key = resolve(filePath);
    const previous = pendingSaves.get(key) ?? Promise.resolve('');
    const current = previous.catch(() => '').then(() =>
        savePortableFileUnlocked(key, bytes, expectedSha256));
    pendingSaves.set(key, current);
    void current.finally(() => {
        if (pendingSaves.get(key) === current) pendingSaves.delete(key);
    }).catch(() => undefined);
    return current;
}

async function savePortableFileUnlocked(
    filePath: string,
    bytes: Uint8Array,
    expectedSha256: string,
): Promise<string> {
    validateBytes(bytes);
    if (!/^[a-f0-9]{64}$/.test(expectedSha256)) {
        throw new Error('Reload this spreadsheet before saving: its source snapshot is missing.');
    }
    const fullPath = resolve(filePath);
    const parent = await localParent(fullPath);
    const initial = await fs.lstat(fullPath);
    if (!initial.isFile() || initial.isSymbolicLink() || initial.nlink !== 1) {
        throw new Error('Portable saves require a regular file without symbolic or hard links.');
    }
    if ((initial.mode & 0o222) === 0) throw new Error('This spreadsheet is read-only.');
    const source = await fs.open(fullPath, constants.O_RDONLY | constants.O_NOFOLLOW);
    let staging: string | undefined;
    let recoveryPath: string | undefined;
    let committed = false;
    try {
        const opened = await source.stat();
        if (opened.dev !== initial.dev || opened.ino !== initial.ino) {
            throw new Error('Spreadsheet was replaced before saving. Reload it first.');
        }
        const previousBytes = await readStable(source);
        if (hash(previousBytes) !== expectedSha256) {
            throw new Error('Spreadsheet changed outside VS Code. Reload it before saving.');
        }
        staging = await fs.mkdtemp(join(parent, '.excel-ai-save-'));
        const stagePath = join(staging, 'workbook');
        await writeExclusive(stagePath, bytes, 0o600);
        const stageInfo = await fs.stat(stagePath);
        if (stageInfo.uid !== initial.uid || stageInfo.gid !== initial.gid) {
            await fs.chown(stagePath, initial.uid, initial.gid);
        }
        await fs.chmod(stagePath, initial.mode & 0o777);

        const backups = join(parent, BACKUP_DIRECTORY);
        await fs.mkdir(backups, { mode: 0o700 }).catch(error => {
            if ((error as NodeJS.ErrnoException).code !== 'EEXIST') throw error;
        });
        const backupInfo = await fs.lstat(backups);
        if (!backupInfo.isDirectory() || backupInfo.isSymbolicLink()
            || (backupInfo.mode & 0o022) !== 0 || await fs.realpath(backups) !== backups) {
            throw new Error('The recovery directory is not a private local directory.');
        }
        recoveryPath = join(backups, `portable-${Date.now()}-${randomUUID()}${extname(fullPath).slice(0, 16)}`);
        await writeExclusive(recoveryPath, previousBytes, 0o600);

        const current = await fs.lstat(fullPath);
        if (current.dev !== initial.dev || current.ino !== initial.ino || current.nlink !== 1
            || current.mode !== initial.mode || current.uid !== initial.uid || current.gid !== initial.gid
            || await fs.realpath(parent) !== parent || hash(await readStable(source)) !== expectedSha256) {
            throw new Error('Spreadsheet changed during saving. Reload it before trying again.');
        }
        await fs.rename(stagePath, fullPath);
        committed = true;
        return hash(bytes);
    } finally {
        await source.close();
        if (staging) await fs.rm(staging, { recursive: true, force: true });
        if (!committed && recoveryPath) await fs.rm(recoveryPath, { force: true });
    }
}
