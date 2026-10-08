import assert from 'node:assert/strict';
import { spawn } from 'node:child_process';
import { randomUUID } from 'node:crypto';
import { mkdirSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';

// Usage: node verify-runtime.mjs <launcher.js | mcp-powerpoint.exe> [--require-powerpoint]
//
// Besides the MCP handshake, this drives one real PowerPoint workflow (create, add a blank
// slide, close) through the given runtime. Listing tools alone cannot catch runtime settings
// that break COM activation in published builds (see issue #103). On machines without
// PowerPoint, `presentation create` must still reach COM activation and report that PowerPoint
// is not installed; any other failure fails the smoke test. Pass --require-powerpoint to fail
// unless the full workflow succeeds against real PowerPoint.
const launcherPath = process.argv[2];
if (!launcherPath) {
  throw new Error('Usage: node verify-runtime.mjs <launcher-path> [--require-powerpoint]');
}
const requirePowerPoint = process.argv.includes('--require-powerpoint');
const [command, args] = launcherPath.toLowerCase().endsWith('.exe')
  ? [launcherPath, []]
  : [process.execPath, [launcherPath]];

const workDir = join(tmpdir(), `powerpoint-mcp-smoke-${randomUUID()}`);
mkdirSync(workDir, { recursive: true });

const child = spawn(command, args, {
  stdio: ['pipe', 'pipe', 'pipe'],
  windowsHide: true
});

let stdout = '';
let stderr = '';
let nextId = 1;
const pending = new Map();

const timeout = setTimeout(() => {
  child.kill();
  console.error(`Timed out waiting for MCP responses.\nstderr: ${stderr}`);
  process.exit(1);
}, 300_000);

child.stderr.on('data', chunk => {
  stderr += chunk.toString('utf8');
});

child.stdout.on('data', chunk => {
  stdout += chunk.toString('utf8');
  const lines = stdout.split(/\r?\n/);
  stdout = lines.pop() ?? '';
  for (const line of lines.filter(value => value.trim())) {
    const message = JSON.parse(line);
    const resolve = pending.get(message.id);
    if (resolve) {
      pending.delete(message.id);
      resolve(message);
    }
  }
});

const exited = new Promise((resolve, reject) => {
  child.once('error', reject);
  child.once('close', resolve);
});

function send(message) {
  child.stdin.write(`${JSON.stringify(message)}\n`);
}

function request(method, params) {
  const id = nextId++;
  const response = new Promise(resolve => pending.set(id, resolve));
  send({ jsonrpc: '2.0', id, method, params });
  return response;
}

async function callTool(name, toolArguments) {
  const response = await request('tools/call', { name, arguments: toolArguments });
  assert.equal(response.error, undefined, `${name} protocol error: ${JSON.stringify(response.error)}`);
  const text = response.result?.content?.find(item => item.type === 'text')?.text;
  assert.ok(text, `${name} returned no text content: ${JSON.stringify(response.result)}`);
  return JSON.parse(text);
}

let workflow;
let failure;
try {
  const initialized = await request('initialize', {
    protocolVersion: '2025-06-18',
    capabilities: {},
    clientInfo: { name: 'powerpoint-mcp-runtime-smoke-test', version: '1.0.0' }
  });
  send({ jsonrpc: '2.0', method: 'notifications/initialized' });
  const tools = await request('tools/list', {});

  assert.equal(initialized?.result?.serverInfo?.name, 'powerpoint-mcp');
  assert.ok(Array.isArray(tools?.result?.tools));
  assert.ok(tools.result.tools.length > 0);
  console.log(
    `MCP handshake succeeded with ${tools.result.tools.length} tools ` +
      `(server ${initialized.result.serverInfo.version}).`
  );

  const created = await callTool('presentation', {
    action: 'create',
    filePath: join(workDir, 'runtime-smoke.pptx')
  });
  if (created.success) {
    assert.ok(created.presentation_session_id, `presentation create returned no presentation_session_id: ${JSON.stringify(created)}`);
    const slide = await callTool('slide', { action: 'add-blank', presentation_session_id: created.presentation_session_id });
    assert.equal(slide.success, true, `slide add-blank failed: ${JSON.stringify(slide)}`);
    const closed = await callTool('presentation', {
      action: 'close',
      presentation_session_id: created.presentation_session_id,
      save: false
    });
    assert.equal(closed.success, true, `presentation close failed: ${JSON.stringify(closed)}`);
    workflow = 'PowerPoint workflow succeeded: presentation create, slide add-blank, presentation close.';
  } else {
    const error = JSON.stringify(created);
    assert.ok(
      !requirePowerPoint && /PowerPoint is not installed/.test(error),
      `presentation create failed: ${error}`
    );
    workflow =
      'PowerPoint is not installed; COM activation was reached and reported it. ' +
      'Real PowerPoint workflow NOT RUN.';
  }
} catch (error) {
  failure = error;
} finally {
  child.stdin.end();
}

const exitCode = await exited;
clearTimeout(timeout);
rmSync(workDir, { recursive: true, force: true });

if (failure) {
  throw failure;
}
assert.equal(exitCode, 0, `Runtime exited with code ${exitCode}.\nstderr: ${stderr}`);
console.log(workflow);