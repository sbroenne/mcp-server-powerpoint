#!/usr/bin/env node

import { createLauncher } from '../lib/launcher.js';

const { main } = createLauncher({
  packageName: '@sbroenne/pptcli',
  commandName: 'pptcli'
});
const exitCode = main();
if (exitCode !== undefined) {
  process.exitCode = exitCode;
}
