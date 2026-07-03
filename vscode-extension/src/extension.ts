import * as vscode from 'vscode';
import * as path from 'path';

/**
 * PowerPoint MCP VS Code Extension
 *
 * Provides an MCP server definition for the PowerPoint MCP server, enabling AI
 * assistants like GitHub Copilot to automate Microsoft PowerPoint through native
 * COM automation.
 *
 * The extension bundles a self-contained executable for the MCP server — no .NET
 * SDK or runtime installation required.
 *
 * Agent Skills are registered via the chatSkills contribution point in package.json.
 */

export async function activate(context: vscode.ExtensionContext) {
	console.log('PowerPoint MCP extension is now active');

	// Register MCP server definition provider
	context.subscriptions.push(
		vscode.lm.registerMcpServerDefinitionProvider('powerpoint-mcp', {
			provideMcpServerDefinitions: async () => {
				const extensionPath = context.extensionPath;
				const mcpServerPath = path.join(extensionPath, 'bin', 'Sbroenne.PowerPointMcp.McpServer.exe');

				return [
					new vscode.McpStdioServerDefinition(
						'powerpoint-mcp',
						mcpServerPath,
						[],
						{
							// Optional environment variables can be added here if needed
						}
					)
				];
			}
		})
	);

	// Show welcome message on first activation
	const hasShownWelcome = context.globalState.get<boolean>('powerpointmcp.hasShownWelcome', false);
	if (!hasShownWelcome) {
		showWelcomeMessage();
		context.globalState.update('powerpointmcp.hasShownWelcome', true);
	}
}

function showWelcomeMessage() {
	const message = 'PowerPoint MCP extension activated! The PowerPoint MCP server is now available for AI assistants.';
	const learnMore = 'Learn More';

	vscode.window.showInformationMessage(message, learnMore).then(selection => {
		if (selection === learnMore) {
			vscode.env.openExternal(vscode.Uri.parse('https://github.com/sbroenne/mcp-server-powerpoint'));
		}
	});
}

export function deactivate() {
	console.log('PowerPoint MCP extension is now deactivated');
}
