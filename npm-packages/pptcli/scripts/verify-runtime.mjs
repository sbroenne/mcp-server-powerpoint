import assert from 'node:assert/strict';
import { spawnSync } from 'node:child_process';
import { randomUUID } from 'node:crypto';
import { mkdirSync, rmSync } from 'node:fs';
import { createRequire } from 'node:module';
import { tmpdir } from 'node:os';
import { join } from 'node:path';

// Usage: node verify-runtime.mjs <launcher.js | pptcli.exe> [--require-powerpoint]
//
// Besides help and launcher parity, this drives one real PowerPoint workflow (session create,
// slide add-blank, session close) through the runtime. Help output alone cannot catch runtime
// settings that break COM activation in published builds (see issue #103). On machines without
// PowerPoint, `session create` must still reach COM activation and report that PowerPoint is not
// installed; any other failure fails the smoke test. Pass --require-powerpoint to fail unless
// the full workflow succeeds against real PowerPoint.
const launcherPath = process.argv[2];
if (!launcherPath) {
  throw new Error('Usage: node verify-runtime.mjs <launcher-path> [--require-powerpoint]');
}
const requirePowerPoint = process.argv.includes('--require-powerpoint');
const directExecutable = launcherPath.toLowerCase().endsWith('.exe');
const runtime = directExecutable
  ? launcherPath
  : createRequire(launcherPath).resolve(`@sbroenne/pptcli-win32-${process.arch}`);
const options = { encoding: 'utf8', timeout: 30_000, windowsHide: true };

for (const args of [['--help'], ['session', '--help'], ['--invalid-npm-smoke-option']]) {
  const direct = spawnSync(runtime, args, options);
  assert.ifError(direct.error);
  assert.equal(direct.signal, null);
  if (args[0] === '--invalid-npm-smoke-option') {
    assert.notEqual(direct.status, 0, 'Invalid arguments must fail.');
  } else {
    assert.equal(direct.status, 0, direct.stderr);
    assert.ok(direct.stdout.trim(), 'CLI help must produce output.');
  }
  if (!directExecutable) {
    const launched = spawnSync(process.execPath, [launcherPath, ...args], options);
    assert.ifError(launched.error);
    assert.equal(launched.signal, null);
    assert.equal(launched.status, direct.status, 'Launcher must preserve exit status.');
    assert.equal(launched.stdout, direct.stdout, 'Launcher must preserve stdout.');
    assert.equal(launched.stderr, direct.stderr, 'Launcher must preserve stderr.');
  }
}

if (!directExecutable) {
  console.log('CLI npm launcher preserves help, subcommands, and failure exit codes.');
}

function run(args) {
  const result = spawnSync(runtime, args, { ...options, timeout: 300_000 });
  assert.ifError(result.error);
  assert.equal(result.signal, null, `pptcli ${args.join(' ')} was terminated.`);
  return { status: result.status, output: `${result.stdout}${result.stderr}`.trim() };
}

const workDir = join(tmpdir(), `pptcli-smoke-${randomUUID()}`);
mkdirSync(workDir, { recursive: true });
let workflow;
try {
  const created = run(['session', 'create', join(workDir, 'runtime-smoke.pptx')]);
  if (created.status === 0) {
    const { sessionId } = JSON.parse(created.output);
    assert.ok(sessionId, `session create returned no sessionId: ${created.output}`);
    const slide = run(['slide', 'add-blank', '--session', sessionId]);
    assert.equal(slide.status, 0, `slide add-blank failed: ${slide.output}`);
    const closed = run(['session', 'close', sessionId]);
    assert.equal(closed.status, 0, `session close failed: ${closed.output}`);
    workflow = 'PowerPoint workflow succeeded: session create, slide add-blank, session close.';
  } else {
    assert.ok(
      !requirePowerPoint && /PowerPoint is not installed/.test(created.output),
      `session create failed: ${created.output}`
    );
    workflow =
      'PowerPoint is not installed; COM activation was reached and reported it. ' +
      'Real PowerPoint workflow NOT RUN.';
  }
} finally {
  if (run(['service', 'stop']).status !== 0) {
    run(['service', 'stop', '--force']);
  }
  rmSync(workDir, { recursive: true, force: true });
}

console.log(workflow);