import type { ExecFileOptions } from 'node:child_process';
import { join } from 'node:path';
import { beforeEach, describe, expect, it, vi } from 'vitest';
import * as vscode from 'vscode';
import { activate } from '../src/extension';
import { noCancellation, output } from './vscode';

const probes = vi.hoisted(() => ({
	access: vi.fn<(path: string, mode?: number) => Promise<void>>(),
	query: vi.fn<(file: string, args: readonly string[], options: ExecFileOptions,
		callback: (error: Error | null, stdout: string, stderr: string) => void) => void>()
}));

vi.mock('node:fs/promises', () => ({ access: probes.access }));
vi.mock('node:child_process', () => ({ execFile: probes.query }));

function createContext() {
	const extensionPath = join('C:', 'fixtures', 'powerpoint-extension');
	return {
		extensionPath,
		extension: {
			id: 'sbroenne.powerpoint-mcp',
			extensionPath,
			extensionUri: vscode.Uri.file(extensionPath),
			isActive: true,
			packageJSON: { version: '1.2.3' },
			exports: undefined,
			extensionKind: vscode.ExtensionKind.UI,
			activate: async () => undefined
		},
		subscriptions: [] as vscode.Disposable[],
		globalState: {
			get: vi.fn().mockReturnValue(true),
			update: vi.fn().mockResolvedValue(undefined),
			keys: () => [],
			setKeysForSync: vi.fn()
		}
	} satisfies Parameters<typeof activate>[0];
}

async function registeredProvider() {
	const context = createContext();
	await activate(context);
	const registration = vi.mocked(vscode.lm.registerMcpServerDefinitionProvider).mock.calls.at(-1);
	expect(registration?.[0]).toBe('powerpoint-mcp');
	if (!registration) throw new Error('The extension did not register its MCP provider.');
	return { context, provider: registration[1] };
}

beforeEach(() => {
	probes.access.mockReset().mockResolvedValue(undefined);
	probes.query.mockReset().mockImplementation((_file, _args, _options, callback) => callback(null, '', ''));
});

describe('MCP registration and launch', () => {
	it('registers the packaged version and bundled command', async () => {
		const { context, provider } = await registeredProvider();
		const definitions = await provider.provideMcpServerDefinitions(noCancellation);
		expect(definitions).toHaveLength(1);
		expect(definitions?.[0]).toMatchObject({
			label: 'powerpoint-mcp',
			command: join(context.extensionPath, 'bin', 'Sbroenne.PowerPointMcp.McpServer.exe'),
			version: '1.2.3'
		});
	});

	it('checks the executable and PowerPoint registration before launch', async () => {
		const { provider } = await registeredProvider();
		const definitions = await provider.provideMcpServerDefinitions(noCancellation);
		const server = definitions?.[0];
		expect(server).toBeDefined();
		expect(provider.resolveMcpServerDefinition).toBeTypeOf('function');
		await provider.resolveMcpServerDefinition?.(server!, noCancellation);
		expect(probes.access).toHaveBeenCalledOnce();
		expect(probes.query).toHaveBeenCalledOnce();
		expect(probes.query.mock.calls[0][1].join(' ')).toContain('PowerPoint.Application');
	});

	it('reports missing PowerPoint without leaking probe details', async () => {
		probes.query.mockImplementation((_file, _args, _options, callback) =>
			callback(Object.assign(new Error('private-fixture'), { code: 2 }), '', ''));
		const { provider } = await registeredProvider();
		const definitions = await provider.provideMcpServerDefinitions(noCancellation);
		await expect(provider.resolveMcpServerDefinition?.(definitions![0], noCancellation))
			.rejects.toThrow(/install.*PowerPoint/i);
		expect(JSON.stringify(output.appendLine.mock.calls)).not.toContain('private-fixture');
	});

	it('maps launch cancellation and disposes the listener', async () => {
		const dispose = vi.fn();
		const token = {
			isCancellationRequested: true,
			onCancellationRequested: vi.fn(() => ({ dispose }))
		};
		const { provider } = await registeredProvider();
		const definitions = await provider.provideMcpServerDefinitions(noCancellation);

		await expect(provider.resolveMcpServerDefinition?.(definitions![0], token))
			.rejects.toBeInstanceOf(vscode.CancellationError);
		expect(dispose).toHaveBeenCalledOnce();
		expect(probes.access).not.toHaveBeenCalled();
		expect(probes.query).not.toHaveBeenCalled();
	});
});
