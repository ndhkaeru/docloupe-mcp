#!/usr/bin/env node
'use strict';

const fs = require('fs');
const https = require('https');
const os = require('os');
const path = require('path');
const { spawn, spawnSync } = require('child_process');

const SERVERS = new Set(['excel', 'md', 'pdf', 'docx', 'pptx', 'csv', 'html', 'text', 'json']);
const RETRYABLE_RENAME_ERRORS = new Set(['EACCES', 'EBUSY', 'EPERM']);
const IN_USE_ERRORS = new Set(['EACCES', 'EBUSY', 'EPERM', 'ETXTBSY']);
const PLATFORM_KEYS = new Set(['win32-x64', 'linux-x64', 'darwin-x64', 'darwin-arm64']);
const RELEASE_TAG_PATTERN = /^v\d+\.\d+\.\d+[0-9A-Za-z.+-]*$/;
const CACHE_FILE_PATTERN = /^[a-z]+-tools(\.exe)?(\.\d+\.\d+)?(\.tmp)?$/;
const CACHE_STALE_MS = 7 * 24 * 60 * 60 * 1000;
const TEMPORARY_STALE_MS = 60 * 60 * 1000;
const DOWNLOAD_IDLE_TIMEOUT_MS = 60 * 1000;
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

function fileSize(filePath) {
  try {
    return fs.statSync(filePath).size;
  } catch {
    return -1;
  }
}

// Another launcher may place the same release asset first; its copy can already be
// running (and locked on Windows), so an identical file is reused instead of replaced.
async function placeDownloadedFile(tmpPath, outputPath, size) {
  if (fileSize(outputPath) === size) {
    fs.rmSync(tmpPath, { force: true });
    return;
  }
  try {
    await renameWithRetry(tmpPath, outputPath);
  } catch (error) {
    fs.rmSync(tmpPath, { force: true });
    if (fileSize(outputPath) === size) return;
    throw error;
  }
  if (process.platform !== 'win32') fs.chmodSync(outputPath, 0o755);
}

// Each download writes its own temporary file and is only renamed into the cache after
// the full Content-Length arrived, so an interrupted download is never cached.
function download(url, outputPath, options = {}, redirects = 0) {
  const get = options.get || https.get;
  return new Promise((resolve, reject) => {
    let settled = false;
    let cleanup = null;
    const fail = (error) => {
      if (settled) return;
      settled = true;
      if (cleanup) cleanup(() => reject(error));
      else reject(error);
    };

    const request = get(url, { headers: { 'User-Agent': 'docloupe-mcp-npm' } }, (response) => {
      if ([301, 302, 303, 307, 308].includes(response.statusCode)) {
        response.resume();
        settled = true;
        if (!response.headers.location || redirects >= 5) {
          reject(new Error(`Too many redirects while downloading ${url}`));
          return;
        }
        download(response.headers.location, outputPath, options, redirects + 1).then(resolve, reject);
        return;
      }
      if (response.statusCode < 200 || response.statusCode >= 300) {
        response.resume();
        fail(new Error(`Download failed (${response.statusCode}): ${url}`));
        return;
      }

      const expected = Number.parseInt(response.headers['content-length'] || '', 10);
      const tmpPath = `${outputPath}.${process.pid}.${Date.now()}.tmp`;
      const file = fs.createWriteStream(tmpPath);
      let received = 0;
      cleanup = (done) => {
        response.destroy();
        const remove = () => fs.rm(tmpPath, { force: true }, () => done());
        if (file.closed) remove();
        else {
          file.once('close', remove);
          file.destroy();
        }
      };

      response.on('data', (chunk) => {
        received += chunk.length;
      });
      response.on('error', fail);
      response.on('close', () => {
        if (!response.complete) fail(new Error(`Download interrupted after ${received} bytes: ${url}`));
      });
      file.on('error', fail);
      file.on('finish', () => {
        file.close((closeError) => {
          if (settled) return;
          if (closeError) {
            fail(closeError);
            return;
          }
          if (Number.isFinite(expected) && received !== expected) {
            fail(new Error(`Download incomplete (${received} of ${expected} bytes): ${url}`));
            return;
          }
          settled = true;
          placeDownloadedFile(tmpPath, outputPath, received).then(resolve, reject);
        });
      });
      response.pipe(file);
    });
    request.setTimeout(DOWNLOAD_IDLE_TIMEOUT_MS, () => {
      request.destroy(new Error(`Download timed out after ${DOWNLOAD_IDLE_TIMEOUT_MS} ms without data: ${url}`));
    });
    request.on('error', fail);
  });
}

function cachePruneEnabled() {
  const value = String(process.env.DOCLOUPE_MCP_CACHE_PRUNE || '').trim().toLowerCase();
  return !['0', 'false', 'no', 'off'].includes(value);
}

function markCacheUsed(root = cacheRoot(), tag = releaseTag(), now = Date.now()) {
  const touchedAt = new Date(now);
  fs.utimesSync(path.join(root, tag), touchedAt, touchedAt);
}

// Lists the files of one release-tag cache directory, or null when it holds anything
// the launcher did not create (such directories are never deleted).
function cacheFiles(tagDirectory) {
  const files = [];
  for (const platformEntry of fs.readdirSync(tagDirectory, { withFileTypes: true })) {
    if (!platformEntry.isDirectory() || !PLATFORM_KEYS.has(platformEntry.name)) return null;
    const platformDirectory = path.join(tagDirectory, platformEntry.name);
    for (const fileEntry of fs.readdirSync(platformDirectory, { withFileTypes: true })) {
      if (!fileEntry.isFile() || !CACHE_FILE_PATTERN.test(fileEntry.name)) return null;
      files.push(path.join(platformDirectory, fileEntry.name));
    }
  }
  return files;
}

// Windows refuses write access to a running executable and Linux reports ETXTBSY.
function isRunning(filePath) {
  if (filePath.endsWith('.tmp')) return false;
  try {
    fs.closeSync(fs.openSync(filePath, 'r+'));
    return false;
  } catch (error) {
    return IN_USE_ERRORS.has(error.code);
  }
}

// Removes release caches unused for a week and stale partial downloads. Directories with
// a running server binary or unexpected content are left alone and retried on a later start.
function pruneCache(root = cacheRoot(), currentTag = releaseTag(), options = {}) {
  const now = options.now || Date.now();
  const maxAgeMs = options.maxAgeMs || CACHE_STALE_MS;
  let entries;
  try {
    entries = fs.readdirSync(root, { withFileTypes: true });
  } catch {
    return;
  }
  for (const entry of entries) {
    if (!entry.isDirectory() || !RELEASE_TAG_PATTERN.test(entry.name)) continue;
    const tagDirectory = path.join(root, entry.name);
    try {
      const files = cacheFiles(tagDirectory);
      if (!files) continue;
      if (entry.name === currentTag) {
        for (const file of files) {
          if (file.endsWith('.tmp') && now - fs.statSync(file).mtimeMs > TEMPORARY_STALE_MS) {
            fs.rmSync(file, { force: true });
          }
        }
        continue;
      }
      if (now - fs.statSync(tagDirectory).mtimeMs <= maxAgeMs) continue;
      if (files.some(isRunning)) continue;
      fs.rmSync(tagDirectory, { recursive: true, force: true });
    } catch {
      // Concurrently removed or locked; a later start retries.
    }
  }
}

function maintainCache() {
  if (!cachePruneEnabled()) return;
  try {
    markCacheUsed();
    pruneCache();
  } catch {
    // Cache maintenance must never block a server start.
  }
}

async function findBinary(server) {
  const override = process.env[envName(server)] || process.env.DOCLOUPE_MCP_BINARY;
  if (override) return override;

  const bundled = path.join(__dirname, '..', 'native', platformKey(), executableName(server));
  if (fs.existsSync(bundled)) return bundled;

  const cached = cachedBinary(server);
  if (!fs.existsSync(cached)) {
    fs.mkdirSync(path.dirname(cached), { recursive: true });
    console.error(`Downloading docloupe ${server}-tools ${releaseTag()} for ${platformKey()}...`);
    await download(assetUrl(server), cached);
  }
  maintainCache();
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
    '  DOCLOUPE_MCP_CACHE_PRUNE=0   (keep release caches unused for over 7 days)',
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
  download,
  executableName,
  markCacheUsed,
  platformKey,
  pruneCache,
  run,
  runAsync,
  signalExitCode,
};

if (require.main === module) void main();
