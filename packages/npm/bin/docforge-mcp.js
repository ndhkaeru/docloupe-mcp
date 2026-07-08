#!/usr/bin/env node
'use strict';

const fs = require('fs');
const path = require('path');
const { spawn } = require('child_process');

const SERVERS = new Set(['excel', 'md', 'pdf', 'docx', 'pptx', 'csv', 'html', 'text', 'json']);

function platformKey() {
  const platform = process.platform;
  const arch = process.arch;
  if (platform === 'win32' && arch === 'x64') return 'win32-x64';
  if (platform === 'linux' && arch === 'x64') return 'linux-x64';
  if (platform === 'darwin' && arch === 'x64') return 'darwin-x64';
  if (platform === 'darwin' && arch === 'arm64') return 'darwin-arm64';
  throw new Error(`Unsupported platform: ${platform}-${arch}. Supported: win32-x64, linux-x64, darwin-x64, darwin-arm64.`);
}

function executableName(server) {
  return `${server}-tools${process.platform === 'win32' ? '.exe' : ''}`;
}

function envName(server) {
  return `DOCFORGE_${server.toUpperCase().replace(/-/g, '_')}_TOOLS_BINARY`;
}

function findBinary(server) {
  const override = process.env[envName(server)] || process.env.DOCFORGE_MCP_BINARY;
  if (override) return override;
  const bundled = path.join(__dirname, '..', 'native', platformKey(), executableName(server));
  return fs.existsSync(bundled) ? bundled : null;
}

function usage() {
  console.error([
    'Usage:',
    '  docforge-mcp <excel|md|pdf|docx|pptx|csv|html|text|json> [server args...]',
    '  docforge-excel-tools [server args...]',
    '',
    'Environment overrides:',
    '  DOCFORGE_EXCEL_TOOLS_BINARY=/path/to/excel-tools',
    '  DOCFORGE_MCP_BINARY=/path/to/server-binary',
  ].join('\n'));
}

function run(server, args) {
  if (!SERVERS.has(server)) {
    usage();
    process.exit(2);
  }

  let binary;
  try {
    binary = findBinary(server);
  } catch (error) {
    console.error(error.message);
    process.exit(1);
  }

  if (!binary) {
    console.error([
      `docforge ${server}-tools binary was not found for this platform.`,
      `Expected bundled path: native/${platformKey()}/${executableName(server)}`,
      `Set ${envName(server)} or DOCFORGE_MCP_BINARY to use a local binary.`,
    ].join('\n'));
    process.exit(1);
  }

  const child = spawn(binary, args, { stdio: 'inherit', windowsHide: true });
  child.on('error', (error) => {
    console.error(error.message);
    process.exit(1);
  });
  child.on('exit', (code, signal) => {
    if (signal) {
      process.kill(process.pid, signal);
      return;
    }
    process.exit(code ?? 0);
  });
}

function main() {
  const invoked = path.basename(process.argv[1] || '').replace(/\.js$/, '');
  const direct = /^docforge-(.+)-tools$/.exec(invoked);
  if (direct) {
    run(direct[1], process.argv.slice(2));
    return;
  }
  const [server, ...args] = process.argv.slice(2);
  run(server, args);
}

module.exports = { run, platformKey, executableName };

if (require.main === module) main();
