#!/usr/bin/env node
'use strict';

const fs = require('fs');
const https = require('https');
const os = require('os');
const path = require('path');
const { spawn, spawnSync } = require('child_process');

const SERVERS = new Set(['excel', 'md', 'pdf', 'docx', 'pptx', 'csv', 'html', 'text', 'json']);
const RETRYABLE_RENAME_ERRORS = new Set(['EACCES', 'EBUSY', 'EPERM']);
const OWNER = 'ndhkaeru';
const REPO = 'docloupe-mcp';
const SIGNAL_NAMES = process.platform === 'win32'
  ? ['SIGINT', 'SIGTERM', 'SIGBREAK']
  : ['SIGHUP', 'SIGINT', 'SIGTERM'];

class LauncherError extends Error {
  constructor(code, phase, message, details = {}, exitCode = 1) {
    super(message);
    this.name = 'LauncherError';
    this.code = code;
    this.phase = phase;
    this.details = details;
    this.exitCode = exitCode;
  }
}

function platformKey() {
  const platform = process.platform;
  const arch = process.arch;
  if (platform === 'win32' && arch === 'x64') return 'win32-x64';
  if (platform === 'linux' && arch === 'x64') return 'linux-x64';
  if (platform === 'darwin' && arch === 'x64') return 'darwin-x64';
  if (platform === 'darwin' && arch === 'arm64') return 'darwin-arm64';
  throw new Error(`Unsupported platform: ${platform}-${arch}. Supported: win32-x64, linux-x64, darwin-x64, darwin-arm64.`);
}

function releasePlatform() {
  return {
    'win32-x64': 'windows-x64',
    'linux-x64': 'linux-x64',
    'darwin-x64': 'macos-x64',
    'darwin-arm64': 'macos-arm64',
  }[platformKey()];
}

function executableName(server) {
  return `${server}-tools${process.platform === 'win32' ? '.exe' : ''}`;
}

function envName(server) {
  return `DOCLOUPE_${server.toUpperCase().replace(/-/g, '_')}_TOOLS_BINARY`;
}

function packageVersion() {
  return require('../package.json').version;
}

function releaseTag() {
  return process.env.DOCLOUPE_MCP_RELEASE_TAG || `v${packageVersion()}`;
}

function cacheRoot() {
  if (process.env.DOCLOUPE_MCP_CACHE_DIR) return process.env.DOCLOUPE_MCP_CACHE_DIR;
  if (process.platform === 'win32' && process.env.LOCALAPPDATA) {
    return path.join(process.env.LOCALAPPDATA, 'docloupe-mcp');
  }
  return path.join(os.homedir(), '.cache', 'docloupe-mcp');
}

function cachedBinary(server) {
  return path.join(cacheRoot(), releaseTag(), platformKey(), executableName(server));
}

function assetUrl(server) {
  const suffix = process.platform === 'win32' ? '.exe' : '';
  const asset = `docloupe-mcp-${server}-tools-${releasePlatform()}${suffix}`;
  return `https://github.com/${OWNER}/${REPO}/releases/download/${releaseTag()}/${asset}`;
}

function renameWithRetry(sourcePath, outputPath, attempt = 0) {
  return new Promise((resolve, reject) => {
    fs.rename(sourcePath, outputPath, (error) => {
      if (!error) {
        resolve();
        return;
      }
      if (!RETRYABLE_RENAME_ERRORS.has(error.code) || attempt >= 9) {
        reject(error);
        return;
      }
      setTimeout(() => {
        renameWithRetry(sourcePath, outputPath, attempt + 1).then(resolve, reject);
      }, 50 * (attempt + 1));
    });
  });
}

function download(url, outputPath, redirects = 0) {
  return new Promise((resolve, reject) => {
    const request = https.get(url, { headers: { 'User-Agent': 'docloupe-mcp-npm' } }, (response) => {
      if ([301, 302, 303, 307, 308].includes(response.statusCode)) {
        response.resume();
        if (!response.headers.location || redirects >= 5) {
          reject(new Error(`Too many redirects while downloading ${url}`));
          return;
        }
        download(response.headers.location, outputPath, redirects + 1).then(resolve, reject);
        return;
      }
      if (response.statusCode < 200 || response.statusCode >= 300) {
        reject(new Error(`Download failed (${response.statusCode}): ${url}`));
        response.resume();
        return;
      }

      const tmpPath = `${outputPath}.tmp`;
      const file = fs.createWriteStream(tmpPath);
      response.pipe(file);
      file.on('finish', () => {
        file.close(() => {
          renameWithRetry(tmpPath, outputPath).then(() => {
            if (process.platform !== 'win32') fs.chmodSync(outputPath, 0o755);
            resolve();
          }, reject);
        });
      });
      file.on('error', (error) => {
        fs.rmSync(tmpPath, { force: true });
        reject(error);
      });
    });
    request.on('error', reject);
  });
}

async function findBinary(server) {
  const override = process.env[envName(server)] || process.env.DOCLOUPE_MCP_BINARY;
  if (override) return override;

  const bundled = path.join(__dirname, '..', 'native', platformKey(), executableName(server));
  if (fs.existsSync(bundled)) return bundled;

  const cached = cachedBinary(server);
  if (fs.existsSync(cached)) return cached;

  fs.mkdirSync(path.dirname(cached), { recursive: true });
  console.error(`Downloading docloupe ${server}-tools ${releaseTag()} for ${platformKey()}...`);
  await download(assetUrl(server), cached);
  return cached;
}

function usage() {
  console.error([
    'Usage:',
    '  docloupe-mcp <excel|md|pdf|docx|pptx|csv|html|text|json> [server args...]',
    '  docloupe-excel-tools [server args...]',
    '',
    'Environment overrides:',
    '  DOCLOUPE_EXCEL_TOOLS_BINARY=/path/to/excel-tools',
    '  DOCLOUPE_MCP_BINARY=/path/to/server-binary',
    '  DOCLOUPE_MCP_CACHE_DIR=/path/to/cache',
    '  DOCLOUPE_MCP_RELEASE_TAG=v1.2.3',
    '  DOCLOUPE_MCP_SHUTDOWN_GRACE_MS=5000',
    '  DOCLOUPE_MCP_TERMINATE_GRACE_MS=2000',
    '  DOCLOUPE_MCP_KILL_GRACE_MS=3000',
  ].join('\n'));
}

function durationFromEnvironment(name, fallback) {
  const value = Number.parseInt(process.env[name] || '', 10);
  if (!Number.isFinite(value) || value < 0 || value > 60000) return fallback;
  return value;
}

function signalExitCode(signalName) {
  const signalNumber = os.constants.signals[signalName];
  return Number.isInteger(signalNumber) ? 128 + signalNumber : 1;
}

function launcherErrorPayload(error) {
  const launcherError = error instanceof LauncherError
    ? error
    : new LauncherError('DOCLOUPE_LAUNCHER_FAILED', 'startup', error.message || String(error));
  return {
    component: 'docloupe-mcp-launcher',
    code: launcherError.code,
    phase: launcherError.phase,
    message: launcherError.message,
    ...launcherError.details,
  };
}

function emitLauncherError(error) {
  console.error(JSON.stringify(launcherErrorPayload(error)));
  if (process.env.DOCLOUPE_MCP_DEBUG === '1' && error.stack) console.error(error.stack);
}

function taskkillTree(processId, force) {
  const systemRoot = process.env.SystemRoot || 'C:\\Windows';
  const taskkill = path.join(systemRoot, 'System32', 'taskkill.exe');
  const command = fs.existsSync(taskkill) ? taskkill : 'taskkill.exe';
  const args = ['/PID', String(processId), '/T'];
  if (force) args.push('/F');
  const result = spawnSync(command, args, {
    stdio: 'ignore',
    windowsHide: true,
    timeout: 5000,
  });
  return result.status === 0;
}

function sendTreeSignal(child, signalName, force = false) {
  if (!child.pid) return true;
  if (process.platform === 'win32') return taskkillTree(child.pid, force);
  try {
    process.kill(-child.pid, force ? 'SIGKILL' : signalName);
    return true;
  } catch (error) {
    if (error.code === 'ESRCH') return true;
    try {
      return child.kill(force ? 'SIGKILL' : signalName);
    } catch {
      return false;
    }
  }
}

function processGroupAlive(processId) {
  if (process.platform === 'win32' || !processId) return false;
  try {
    process.kill(-processId, 0);
    return true;
  } catch (error) {
    return error.code !== 'ESRCH';
  }
}

function waitForProcessGroupExit(processId, timeoutMs) {
  const deadline = Date.now() + timeoutMs;
  return new Promise((resolve) => {
    const poll = () => {
      if (!processGroupAlive(processId)) {
        resolve(true);
        return;
      }
      if (Date.now() >= deadline) {
        resolve(false);
        return;
      }
      setTimeout(poll, 25);
    };
    poll();
  });
}

function waitWithTimeout(promise, timeoutMs) {
  let timer;
  const timeout = new Promise((resolve) => {
    timer = setTimeout(() => resolve({ completed: false }), timeoutMs);
  });
  return Promise.race([
    promise.then((value) => ({ completed: true, value })),
    timeout,
  ]).finally(() => clearTimeout(timer));
}

function closeChildInput(child) {
  if (!child.stdin || child.stdin.destroyed || child.stdin.writableEnded) return;
  child.stdin.end();
}

async function ensureTreeStopped(child, terminateGraceMs, killGraceMs) {
  if (process.platform === 'win32' || !processGroupAlive(child.pid)) return true;
  sendTreeSignal(child, 'SIGTERM', false);
  if (await waitForProcessGroupExit(child.pid, terminateGraceMs)) return true;
  sendTreeSignal(child, 'SIGKILL', true);
  return waitForProcessGroupExit(child.pid, killGraceMs);
}

async function shutdownChild(child, closePromise, trigger, timings) {
  closeChildInput(child);
  if (trigger.signal) sendTreeSignal(child, trigger.signal, false);

  let closed = await waitWithTimeout(closePromise, timings.shutdownGraceMs);
  if (!closed.completed) {
    sendTreeSignal(child, 'SIGTERM', false);
    closed = await waitWithTimeout(closePromise, timings.terminateGraceMs);
  }
  if (!closed.completed) {
    sendTreeSignal(child, 'SIGKILL', true);
    closed = await waitWithTimeout(closePromise, timings.killGraceMs);
  }
  if (!closed.completed) {
    child.unref();
    throw new LauncherError(
      'DOCLOUPE_LAUNCHER_SHUTDOWN_TIMEOUT',
      'shutdown',
      'Child MCP process did not exit after graceful, terminate, and kill stages.',
      {
        pid: child.pid,
        trigger: trigger.kind,
        signal: trigger.signal || null,
        shutdown_grace_ms: timings.shutdownGraceMs,
        terminate_grace_ms: timings.terminateGraceMs,
        kill_grace_ms: timings.killGraceMs,
      },
    );
  }

  const treeStopped = await ensureTreeStopped(
    child,
    timings.terminateGraceMs,
    timings.killGraceMs,
  );
  if (!treeStopped) {
    throw new LauncherError(
      'DOCLOUPE_LAUNCHER_TREE_STILL_RUNNING',
      'shutdown',
      'Child MCP process exited but its process group is still running.',
      { pid: child.pid, trigger: trigger.kind },
    );
  }
  return closed.value;
}

function childExitCode(closeResult, requestedSignal) {
  if (requestedSignal) return signalExitCode(requestedSignal);
  if (Number.isInteger(closeResult.code)) return closeResult.code;
  if (closeResult.signal) return signalExitCode(closeResult.signal);
  return 1;
}

async function runAsync(server, args) {
  if (!SERVERS.has(server)) {
    usage();
    throw new LauncherError(
      'DOCLOUPE_LAUNCHER_INVALID_SERVER',
      'startup',
      `Unsupported DocLoupe server: ${server || '<missing>'}`,
      { server: server || null },
      2,
    );
  }

  const binary = await findBinary(server);
  const child = spawn(binary, args, {
    stdio: ['pipe', 'inherit', 'inherit'],
    windowsHide: true,
    detached: process.platform !== 'win32',
  });
  const timings = {
    shutdownGraceMs: durationFromEnvironment('DOCLOUPE_MCP_SHUTDOWN_GRACE_MS', 5000),
    terminateGraceMs: durationFromEnvironment('DOCLOUPE_MCP_TERMINATE_GRACE_MS', 2000),
    killGraceMs: durationFromEnvironment('DOCLOUPE_MCP_KILL_GRACE_MS', 3000),
  };

  const closePromise = new Promise((resolve, reject) => {
    child.once('error', (error) => {
      reject(new LauncherError(
        'DOCLOUPE_LAUNCHER_SPAWN_FAILED',
        'startup',
        error.message,
        { binary, server, error_code: error.code || null },
      ));
    });
    child.once('close', (code, signalName) => resolve({ code, signal: signalName }));
  });

  let resolveTrigger;
  let triggerRequested = false;
  const triggerPromise = new Promise((resolve) => {
    resolveTrigger = resolve;
  });
  const requestShutdown = (trigger) => {
    if (triggerRequested) return;
    triggerRequested = true;
    resolveTrigger(trigger);
  };

  const onStdinEnd = () => requestShutdown({ kind: 'stdin_eof', signal: null });
  const onStdinError = (error) => requestShutdown({ kind: 'stdin_error', signal: null, error });
  const onChildStdinError = (error) => {
    if (error.code !== 'EPIPE' && error.code !== 'ERR_STREAM_DESTROYED') {
      requestShutdown({ kind: 'child_stdin_error', signal: null, error });
    }
  };
  const signalHandlers = new Map();
  for (const signalName of SIGNAL_NAMES) {
    const handler = () => requestShutdown({ kind: 'signal', signal: signalName });
    signalHandlers.set(signalName, handler);
    process.on(signalName, handler);
  }

  process.stdin.on('end', onStdinEnd);
  process.stdin.on('close', onStdinEnd);
  process.stdin.on('error', onStdinError);
  child.stdin.on('error', onChildStdinError);
  process.stdin.pipe(child.stdin);
  if (process.stdin.readableEnded || process.stdin.destroyed) queueMicrotask(onStdinEnd);

  try {
    const first = await Promise.race([
      closePromise.then((value) => ({ kind: 'child_exit', value })),
      triggerPromise.then((value) => ({ kind: 'shutdown', value })),
    ]);
    let closeResult;
    let requestedSignal = null;
    if (first.kind === 'child_exit') {
      closeResult = first.value;
      const treeStopped = await ensureTreeStopped(
        child,
        timings.terminateGraceMs,
        timings.killGraceMs,
      );
      if (!treeStopped) {
        throw new LauncherError(
          'DOCLOUPE_LAUNCHER_TREE_STILL_RUNNING',
          'shutdown',
          'Child MCP process exited but its process group is still running.',
          { pid: child.pid, trigger: 'child_exit' },
        );
      }
    } else {
      requestedSignal = first.value.signal;
      closeResult = await shutdownChild(child, closePromise, first.value, timings);
      if (first.value.error) {
        throw new LauncherError(
          'DOCLOUPE_LAUNCHER_STDIN_FAILED',
          'shutdown',
          first.value.error.message,
          { pid: child.pid, trigger: first.value.kind },
        );
      }
    }
    return {
      exitCode: childExitCode(closeResult, requestedSignal),
      childCode: closeResult.code,
      childSignal: closeResult.signal,
      requestedSignal,
    };
  } finally {
    process.stdin.unpipe(child.stdin);
    process.stdin.removeListener('end', onStdinEnd);
    process.stdin.removeListener('close', onStdinEnd);
    process.stdin.removeListener('error', onStdinError);
    child.stdin.removeListener('error', onChildStdinError);
    for (const [signalName, handler] of signalHandlers) {
      process.removeListener(signalName, handler);
    }
    process.stdin.pause();
  }
}

async function run(server, args) {
  try {
    const result = await runAsync(server, args);
    process.exitCode = result.exitCode;
    return result;
  } catch (error) {
    emitLauncherError(error);
    process.exitCode = error instanceof LauncherError ? error.exitCode : 1;
    return null;
  }
}

function main() {
  const invoked = path.basename(process.argv[1] || '').replace(/\.js$/, '');
  const direct = /^docloupe-(.+)-tools$/.exec(invoked);
  if (direct) return run(direct[1], process.argv.slice(2));
  const [server, ...args] = process.argv.slice(2);
  return run(server, args);
}

module.exports = {
  LauncherError,
  childExitCode,
  executableName,
  platformKey,
  run,
  runAsync,
  signalExitCode,
};

if (require.main === module) void main();
