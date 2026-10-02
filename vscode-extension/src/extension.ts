import * as vscode from 'vscode';
import { join } from 'node:path';
import { checkLaunchPrerequisites, LaunchSetupError } from './prerequisites';

const userGuideUrl = 'https://powerpointmcpserver.dev/';

export async function activate(
	context: Pick<vscode.ExtensionContext, 'extension' | 'extensionPath' | 'globalState' | 'subscriptions'>
) {
	const output = vscode.window.createOutputChannel('PowerPointMcp');
	context.subscriptions.push(output);

	const version: unknown = context.extension.packageJSON.version;
	if (typeof version !== 'string' || !/^\d+\.\d+\.\d+(?:-[A-Za-z0-9.-]+)?$/.test(version)) {
		const message = 'PowerPointMcp package version is invalid. Reinstall the extension.';
		output.appendLine(message);
		throw new Error(message);
	}

	const executable = join(context.extensionPath, 'bin', 'Sbroenne.PowerPointMcp.McpServer.exe');
	context.subscriptions.push(
		vscode.lm.registerMcpServerDefinitionProvider('powerpoint-mcp', {
			provideMcpServerDefinitions: async () => [
				new vscode.McpStdioServerDefinition('powerpoint-mcp', executable, [], {}, version)
			],
			resolveMcpServerDefinition: async (server, token) => {
				const controller = new AbortController();
				const cancellation = token.onCancellationRequested(() => controller.abort());
				if (token.isCancellationRequested) {
					controller.abort();
				}
				try {
					await checkLaunchPrerequisites(executable, controller.signal);
					output.appendLine('Launch prerequisites verified. VS Code manages server startup and approvals.');
					return server;
				} catch (error) {
					if (controller.signal.aborted) {
						throw new vscode.CancellationError();
					}
					if (error instanceof LaunchSetupError) {
						output.appendLine(error.message);
						void showSetupError(error.message, output);
					}
					throw error;
				} finally {
					cancellation.dispose();
				}
			}
		})
	);
	output.appendLine(
		`Registered bundled MCP server version ${version}. Server logs: MCP: List Servers > powerpoint-mcp > Show Output.`
	);

	const hasShownWelcome = context.globalState.get<boolean>('powerpointmcp.hasShownWelcome', false);
	if (!hasShownWelcome) {
		void showWelcomeMessage(output);
		try {
			await context.globalState.update('powerpointmcp.hasShownWelcome', true);
		} catch (error) {
			const detail = error instanceof Error ? error.message : String(error);
			output.appendLine(`Could not save the welcome preference. Getting-started help may appear again. ${detail}`);
		}
	}
}

async function showWelcomeMessage(output: vscode.OutputChannel) {
	try {
		const selection = await vscode.window.showInformationMessage(
			'PowerPointMcp bundles real PowerPoint automation. Send a presentation request in Copilot Chat with tool support; VS Code starts powerpoint-mcp automatically when needed. Approve server or tool use if prompted.',
			'Getting Started'
		);
		if (selection === 'Getting Started' && !await vscode.env.openExternal(vscode.Uri.parse(userGuideUrl))) {
			output.appendLine(`Could not open the user guide. Visit ${userGuideUrl}`);
		}
	} catch {
		output.appendLine(`Could not display getting-started help. Visit ${userGuideUrl}`);
	}
}

async function showSetupError(message: string, output: vscode.OutputChannel) {
	try {
		if (await vscode.window.showErrorMessage(message, 'Show Setup Output') === 'Show Setup Output') {
			output.show();
		}
	} catch {
		output.appendLine('Could not display the setup notification. See the setup error above.');
	}
}
