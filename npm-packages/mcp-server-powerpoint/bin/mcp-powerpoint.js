#!/usr/bin/env node

import { createLauncher } from '../lib/launcher.js';

const { main } = createLauncher({
  packageName: '@sbroenne/mcp-server-powerpoint',
  commandName: 'mcp-powerpoint'
});

const exitCode = main();
if (exitCode !== undefined) {
  process.exitCode = exitCode;
}
