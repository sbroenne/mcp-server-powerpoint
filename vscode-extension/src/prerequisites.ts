import { execFile } from 'node:child_process';
import { constants } from 'node:fs';
import { access } from 'node:fs/promises';
import { join } from 'node:path';
import { promisify } from 'node:util';

const run = promisify(execFile);
const registrationCheck = `
$ErrorActionPreference = 'Stop'
$key = [Microsoft.Win32.Registry]::ClassesRoot.OpenSubKey('PowerPoint.Application\\CLSID')
if ($null -eq $key) { exit 2 }
try {
    $clsid = [guid]::Empty
    if (-not [guid]::TryParse([string]$key.GetValue(''), [ref]$clsid)) { exit 2 }
} finally { $key.Dispose() }
`;

export class LaunchSetupError extends Error {}

function checkCancellation(signal: AbortSignal) {
	if (signal.aborted) {
		throw new DOMException('PowerPointMcp startup canceled.', 'AbortError');
	}
}

export async function checkLaunchPrerequisites(executable: string, signal: AbortSignal) {
	checkCancellation(signal);
	if (process.platform !== 'win32') {
		throw new LaunchSetupError('PowerPointMcp requires a local Windows desktop with Microsoft PowerPoint.');
	}
	try {
		await access(executable, constants.R_OK);
	} catch {
		throw new LaunchSetupError(
			'The bundled PowerPointMcp server is missing or unreadable. Check access permissions or reinstall the extension.'
		);
	}
	checkCancellation(signal);

	const windowsDirectory = process.env.SystemRoot;
	if (!windowsDirectory) {
		throw new LaunchSetupError(
			'Could not check PowerPoint registration: the Windows system directory is not configured.'
		);
	}
	const powershell = join(windowsDirectory, 'System32', 'WindowsPowerShell', 'v1.0', 'powershell.exe');
	try {
		await run(powershell, ['-NoProfile', '-NonInteractive', '-Command', registrationCheck], {
			timeout: 10000,
			windowsHide: true,
			maxBuffer: 65536,
			signal
		});
	} catch (error) {
		checkCancellation(signal);
		if (error instanceof Error && 'code' in error && error.code === 2) {
			throw new LaunchSetupError(
				'Microsoft PowerPoint is not registered. Install or repair desktop PowerPoint, open it once, and retry.'
			);
		}
		if (error instanceof Error && 'killed' in error && error.killed === true) {
			throw new LaunchSetupError(
				'Checking PowerPoint registration timed out. Verify Windows PowerShell and desktop PowerPoint work, then retry.'
			);
		}
		throw new LaunchSetupError(
			'Could not check PowerPoint registration. Verify Windows PowerShell works, then retry or repair Microsoft Office.'
		);
	}
	checkCancellation(signal);
}
