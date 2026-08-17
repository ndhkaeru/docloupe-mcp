'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('fs');
const os = require('os');
const path = require('path');
const { spawn, spawnSync } = require('child_process');

const PACKAGE_ROOT = path.resolve(__dirname, '..');
const LAUNCHER = path.join(PACKAGE_ROOT, 'bin', 'docloupe-mcp.js');
const HELPER = path.join(__dirname, 'fixtures', 'launcher-child.js');

function launcherEnvironment(overrides = {}) {
  return {
    ...process.env,
    DOCLOUPE_MCP_BINARY: process.execPath,
    DOCLOUPE_MCP_SHUTDOWN_GRACE_MS: '100',
    DOCLOUPE_MCP_TERMINATE_GRACE_MS: '100',
    DOCLOUPE_MCP_KILL_GRACE_MS: '2000',
    ...overrides,
  };
}

function startLauncher(args, environment = {}) {
  return spawn(process.execPath, [LAUNCHER, 'excel', HELPER, ...args], {
    cwd: PACKAGE_ROOT,
    env: launcherEnvironment(environment),
    stdio: ['pipe', 'pipe', 'pipe'],
    windowsHide: true,
  });
}

function collect(stream) {
  const chunks = [];
  stream.on('data', (chunk) => chunks.push(Buffer.from(chunk)));
  return () => Buffer.concat(chunks).toString('utf8');
}

function waitForClose(child, timeoutMs = 8000) {
  return new Promise((resolve, reject) => {
    const timer = setTimeout(() => {
      child.kill('SIGKILL');
      reject(new Error(`Timed out waiting for launcher PID ${child.pid}`));
    }, timeoutMs);
    child.once('error', (error) => {
      clearTimeout(timer);
      reject(error);
    });
    child.once('close', (code, signal) => {
      clearTimeout(timer);
      resolve({ code, signal });
    });
  });
}

async function waitForFile(filePath, timeoutMs = 5000) {
  const deadline = Date.now() + timeoutMs;
  while (Date.now() < deadline) {
    if (fs.existsSync(filePath)) return;
    await new Promise((resolve) => setTimeout(resolve, 25));
  }
  throw new Error(`Timed out waiting for ${filePath}`);
}

function pidExists(processId) {
  try {
    process.kill(processId, 0);
    return true;
  } catch {
    return false;
  }
}

async function waitForPidsToStop(processIds, timeoutMs = 5000) {
  const deadline = Date.now() + timeoutMs;
  while (Date.now() < deadline) {
    if (!processIds.some(pidExists)) return true;
    await new Promise((resolve) => setTimeout(resolve, 50));
  }
  return !processIds.some(pidExists);
}

function forceStop(processIds) {
  for (const processId of processIds) {
    if (!pidExists(processId)) continue;
    if (process.platform === 'win32') {
      spawnSync('taskkill.exe', ['/PID', String(processId), '/T', '/F'], { stdio: 'ignore' });
    } else {
      try {
        process.kill(processId, 'SIGKILL');
      } catch {
        continue;
      }
    }
  }
}

test('launcher forwards stdin EOF and child exit code', { timeout: 10000 }, async () => {
  const temporaryDirectory = fs.mkdtempSync(path.join(os.tmpdir(), 'docloupe-launcher-'));
  const outputPath = path.join(temporaryDirectory, 'stdin.bin');
  const launcher = startLauncher(['echo', outputPath]);
  const stderrText = collect(launcher.stderr);

  launcher.stdin.end(Buffer.from('mcp-input'));
  const result = await waitForClose(launcher);

  assert.equal(result.code, 7, stderrText());
  assert.equal(fs.readFileSync(outputPath, 'utf8'), 'mcp-input');
  fs.rmSync(temporaryDirectory, { recursive: true, force: true });
});

test('launcher escalates EOF shutdown and kills the child tree', { timeout: 15000 }, async () => {
  const temporaryDirectory = fs.mkdtempSync(path.join(os.tmpdir(), 'docloupe-launcher-'));
  const pidPath = path.join(temporaryDirectory, 'pids.json');
  const launcher = startLauncher(['tree', pidPath]);
  const stderrText = collect(launcher.stderr);
  let processIds = [];

  try {
    await waitForFile(pidPath);
    processIds = Object.values(JSON.parse(fs.readFileSync(pidPath, 'utf8')));
    launcher.stdin.end();
    const result = await waitForClose(launcher);

    assert.notEqual(result.code, 0, stderrText());
    assert.equal(await waitForPidsToStop(processIds), true);
  } finally {
    forceStop(processIds);
    fs.rmSync(temporaryDirectory, { recursive: true, force: true });
  }
});

test('launcher reports structured startup errors', { timeout: 10000 }, async () => {
  const missingBinary = path.join(os.tmpdir(), `missing-docloupe-${process.pid}.exe`);
  const launcher = spawn(process.execPath, [LAUNCHER, 'excel'], {
    cwd: PACKAGE_ROOT,
    env: launcherEnvironment({ DOCLOUPE_MCP_BINARY: missingBinary }),
    stdio: ['pipe', 'pipe', 'pipe'],
    windowsHide: true,
  });
  const stderrText = collect(launcher.stderr);
  launcher.stdin.end();
  const result = await waitForClose(launcher);
  const payloadLine = stderrText().trim().split(/\r?\n/).find((line) => line.startsWith('{'));

  assert.equal(result.code, 1);
  assert.ok(payloadLine, stderrText());
  const payload = JSON.parse(payloadLine);
  assert.equal(payload.component, 'docloupe-mcp-launcher');
  assert.equal(payload.code, 'DOCLOUPE_LAUNCHER_SPAWN_FAILED');
  assert.equal(payload.phase, 'startup');
});

test(
  'launcher forwards POSIX termination signals',
  { timeout: 10000, skip: process.platform === 'win32' },
  async () => {
    const temporaryDirectory = fs.mkdtempSync(path.join(os.tmpdir(), 'docloupe-launcher-'));
    const readyPath = path.join(temporaryDirectory, 'ready.txt');
    const signalPath = path.join(temporaryDirectory, 'signal.txt');
    const launcher = startLauncher(['signal', readyPath, signalPath]);
    const stderrText = collect(launcher.stderr);

    try {
      await waitForFile(readyPath);
      launcher.kill('SIGTERM');
      const result = await waitForClose(launcher);

      assert.equal(result.code, 143, stderrText());
      assert.equal(fs.readFileSync(signalPath, 'utf8'), 'SIGTERM');
    } finally {
      if (pidExists(launcher.pid)) launcher.kill('SIGKILL');
      fs.rmSync(temporaryDirectory, { recursive: true, force: true });
    }
  },
);
