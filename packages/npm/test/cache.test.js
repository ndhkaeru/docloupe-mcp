'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const crypto = require('crypto');
const fs = require('fs');
const http = require('http');
const os = require('os');
const path = require('path');
const { spawn, spawnSync } = require('child_process');

const PACKAGE_ROOT = path.resolve(__dirname, '..');
const launcher = require('../bin/docloupe-mcp.js');

const DAY_MS = 24 * 60 * 60 * 1000;
const standIn = process.platform === 'win32'
  ? path.join(process.env.SystemRoot || 'C:\\Windows', 'System32', 'sort.exe')
  : '/bin/cat';

function temporaryDirectory(t) {
  const directory = fs.mkdtempSync(path.join(os.tmpdir(), 'docloupe-cache-'));
  t.after(() => fs.rmSync(directory, { recursive: true, force: true, maxRetries: 5, retryDelay: 100 }));
  return directory;
}

function serve(t, handler) {
  return new Promise((resolve) => {
    const server = http.createServer(handler);
    server.listen(0, '127.0.0.1', () => {
      t.after(() => new Promise((done) => {
        if (server.closeAllConnections) server.closeAllConnections();
        server.close(done);
      }));
      resolve(`http://127.0.0.1:${server.address().port}/asset`);
    });
  });
}

function sendSlowly(response, payload, chunks, delayMs) {
  response.writeHead(200, { 'Content-Length': payload.length });
  const size = Math.ceil(payload.length / chunks);
  let offset = 0;
  const next = () => {
    if (offset >= payload.length) {
      response.end();
      return;
    }
    response.write(payload.subarray(offset, offset + size));
    offset += size;
    setTimeout(next, delayMs);
  };
  next();
}

function leftovers(directory) {
  return fs.readdirSync(directory).filter((name) => name.endsWith('.tmp'));
}

function put(file, content = 'x') {
  fs.mkdirSync(path.dirname(file), { recursive: true });
  fs.writeFileSync(file, content);
}

function setAge(file, ageMs, now = Date.now()) {
  const time = new Date(now - ageMs);
  fs.utimesSync(file, time, time);
}

function waitForExit(child, timeoutMs = 10000) {
  return new Promise((resolve, reject) => {
    if (child.exitCode !== null) {
      resolve(child.exitCode);
      return;
    }
    const timer = setTimeout(() => reject(new Error(`PID ${child.pid} did not exit`)), timeoutMs);
    child.once('exit', (code) => {
      clearTimeout(timer);
      resolve(code);
    });
  });
}

async function waitUntilRunning(executable, timeoutMs = 5000) {
  const deadline = Date.now() + timeoutMs;
  while (Date.now() < deadline) {
    try {
      fs.closeSync(fs.openSync(executable, 'r+'));
    } catch (error) {
      if (['EBUSY', 'EPERM', 'EACCES', 'ETXTBSY'].includes(error.code)) return;
    }
    await new Promise((resolve) => setTimeout(resolve, 50));
  }
  throw new Error(`${executable} never started`);
}

test('download writes the complete asset without temporary leftovers', async (t) => {
  const directory = temporaryDirectory(t);
  const payload = crypto.randomBytes(256 * 1024);
  const url = await serve(t, (request, response) => sendSlowly(response, payload, 4, 5));
  const output = path.join(directory, 'excel-tools.exe');

  await launcher.download(url, output, { get: http.get });

  assert.ok(fs.readFileSync(output).equals(payload));
  assert.deepEqual(leftovers(directory), []);
});

test('download rejects an interrupted response and caches nothing', async (t) => {
  const directory = temporaryDirectory(t);
  const url = await serve(t, (request, response) => {
    response.writeHead(200, { 'Content-Length': 100000 });
    response.write(Buffer.alloc(40000));
    setTimeout(() => response.socket.destroy(), 20);
  });
  const output = path.join(directory, 'excel-tools.exe');

  await assert.rejects(launcher.download(url, output, { get: http.get }), /interrupted|incomplete|aborted|socket hang up/i);

  assert.equal(fs.existsSync(output), false);
  assert.deepEqual(leftovers(directory), []);
});

test('concurrent downloads of one asset both succeed with one intact file', async (t) => {
  const directory = temporaryDirectory(t);
  const payload = crypto.randomBytes(512 * 1024);
  const url = await serve(t, (request, response) => sendSlowly(response, payload, 8, 10));
  const output = path.join(directory, 'md-tools.exe');

  await Promise.all([
    launcher.download(url, output, { get: http.get }),
    launcher.download(url, output, { get: http.get }),
  ]);

  assert.ok(fs.readFileSync(output).equals(payload));
  assert.deepEqual(leftovers(directory), []);
});

test('pruneCache removes only stale release caches created by the launcher', (t) => {
  const root = temporaryDirectory(t);
  const now = Date.now();
  const stale = path.join(root, 'v1.0.0');
  put(path.join(stale, 'win32-x64', 'excel-tools.exe'));
  put(path.join(stale, 'win32-x64', 'md-tools.exe.tmp'));
  const recent = path.join(root, 'v1.0.1');
  put(path.join(recent, 'linux-x64', 'excel-tools'));
  const current = path.join(root, 'v1.1.0');
  put(path.join(current, 'win32-x64', 'excel-tools.exe'));
  const staleTemporary = path.join(current, 'win32-x64', 'md-tools.exe.tmp');
  const freshTemporary = path.join(current, 'win32-x64', 'md-tools.exe.123.456.tmp');
  put(staleTemporary);
  put(freshTemporary);
  setAge(staleTemporary, 2 * 60 * 60 * 1000, now);
  const foreignName = path.join(root, 'notes');
  put(path.join(foreignName, 'keep.txt'));
  const foreignFile = path.join(root, 'v0.9.0');
  put(path.join(foreignFile, 'win32-x64', 'excel-tools.exe'));
  put(path.join(foreignFile, 'win32-x64', 'readme.txt'));
  const foreignPlatform = path.join(root, 'v0.8.0');
  put(path.join(foreignPlatform, 'solaris-sparc', 'excel-tools'));
  for (const directory of [stale, current, foreignName, foreignFile, foreignPlatform]) setAge(directory, 30 * DAY_MS, now);
  setAge(recent, DAY_MS, now);

  launcher.pruneCache(root, 'v1.1.0', { now });

  assert.equal(fs.existsSync(stale), false, 'stale cache is removed');
  assert.equal(fs.existsSync(recent), true, 'cache used within a week is kept');
  assert.equal(fs.existsSync(path.join(current, 'win32-x64', 'excel-tools.exe')), true, 'current release is kept');
  assert.equal(fs.existsSync(staleTemporary), false, 'stale partial download is removed');
  assert.equal(fs.existsSync(freshTemporary), true, 'partial download of a concurrent launcher is kept');
  assert.equal(fs.existsSync(foreignName), true, 'non-release directory is ignored');
  assert.equal(fs.existsSync(path.join(foreignFile, 'win32-x64', 'readme.txt')), true, 'directory with foreign files is ignored');
  assert.equal(fs.existsSync(foreignPlatform), true, 'directory with unknown platform is ignored');

  launcher.pruneCache(path.join(root, 'missing'), 'v1.1.0', { now });
});

test(
  'pruneCache keeps a stale cache while its server binary is running',
  { skip: process.platform === 'darwin' || !fs.existsSync(standIn), timeout: 20000 },
  async (t) => {
    const root = temporaryDirectory(t);
    const executable = path.join(root, 'v1.0.0', launcher.platformKey(), launcher.executableName('excel'));
    fs.mkdirSync(path.dirname(executable), { recursive: true });
    fs.copyFileSync(standIn, executable);
    fs.chmodSync(executable, 0o755);
    const child = spawn(executable, [], { stdio: ['pipe', 'ignore', 'ignore'], windowsHide: true });
    t.after(() => {
      if (child.exitCode === null) child.kill('SIGKILL');
    });
    await waitUntilRunning(executable);
    setAge(path.join(root, 'v1.0.0'), 30 * DAY_MS);

    launcher.pruneCache(root, 'v2.0.0');
    assert.equal(fs.existsSync(executable), true, 'running binary is kept');

    child.stdin.end();
    await waitForExit(child);
    setAge(path.join(root, 'v1.0.0'), 30 * DAY_MS);
    launcher.pruneCache(root, 'v2.0.0');
    assert.equal(fs.existsSync(path.join(root, 'v1.0.0')), false, 'cache is removed once the server exits');
  },
);

function npm(args, cwd, environment) {
  const quote = (argument) => (/[\s"&|<>^]/.test(argument) ? `"${argument}"` : argument);
  const result = spawnSync(['npm', ...args].map(quote).join(' '), { cwd, env: environment, encoding: 'utf8', shell: true });
  return { status: result.status, output: `${result.stdout || ''}${result.stderr || ''}` };
}

test(
  'npm upgrade while a server runs does not hit locked files on Windows',
  { skip: process.platform !== 'win32', timeout: 180000 },
  async (t) => {
    const root = temporaryDirectory(t);
    const manifest = require('../package.json');
    const cacheDirectory = path.join(root, 'cache');
    const environment = {
      ...process.env,
      npm_config_cache: path.join(root, 'npm-cache'),
      npm_config_update_notifier: 'false',
    };
    const versions = ['0.0.0-upgrade.1', '0.0.0-upgrade.2'];
    const tarballs = versions.map((version) => {
      const stage = path.join(root, 'stage', version);
      fs.mkdirSync(path.join(stage, 'bin'), { recursive: true });
      for (const name of fs.readdirSync(path.join(PACKAGE_ROOT, 'bin'))) {
        fs.copyFileSync(path.join(PACKAGE_ROOT, 'bin', name), path.join(stage, 'bin', name));
      }
      fs.writeFileSync(path.join(stage, 'package.json'), JSON.stringify({ ...manifest, version, scripts: {} }, null, 2));
      const cached = path.join(cacheDirectory, `v${version}`, launcher.platformKey(), launcher.executableName('excel'));
      fs.mkdirSync(path.dirname(cached), { recursive: true });
      fs.copyFileSync(standIn, cached);
      fs.mkdirSync(path.join(root, 'artifacts'), { recursive: true });
      const packed = npm(['pack', stage, '--pack-destination', path.join(root, 'artifacts'), '--json'], root, environment);
      assert.equal(packed.status, 0, packed.output);
      return path.join(root, 'artifacts', JSON.parse(packed.output.slice(packed.output.indexOf('[')))[0].filename);
    });
    const installRoot = path.join(root, 'install');
    const install = (tarball) => {
      fs.mkdirSync(installRoot, { recursive: true });
      fs.writeFileSync(path.join(installRoot, 'package.json'), JSON.stringify({
        private: true,
        dependencies: { [manifest.name]: `file:${tarball.replace(/\\/g, '/')}` },
      }));
      return npm(['install', '--ignore-scripts', '--no-audit', '--no-fund', '--package-lock=false'], installRoot, environment);
    };

    const first = install(tarballs[0]);
    assert.equal(first.status, 0, first.output);
    const server = spawn(process.execPath, [
      path.join(installRoot, 'node_modules', '@ndhkaeru', 'docloupe-mcp', 'bin', 'docloupe-excel-tools.js'),
    ], {
      cwd: installRoot,
      env: { ...process.env, DOCLOUPE_MCP_CACHE_DIR: cacheDirectory, DOCLOUPE_MCP_BINARY: '', DOCLOUPE_EXCEL_TOOLS_BINARY: '' },
      stdio: ['pipe', 'ignore', 'pipe'],
      windowsHide: true,
    });
    t.after(() => {
      if (server.exitCode === null) spawnSync('taskkill.exe', ['/PID', String(server.pid), '/T', '/F'], { stdio: 'ignore' });
    });
    await waitUntilRunning(path.join(cacheDirectory, `v${versions[0]}`, launcher.platformKey(), launcher.executableName('excel')));

    const upgrade = install(tarballs[1]);
    const scope = path.join(installRoot, 'node_modules', '@ndhkaeru');
    const installed = JSON.parse(fs.readFileSync(path.join(scope, 'docloupe-mcp', 'package.json'), 'utf8')).version;

    assert.equal(upgrade.status, 0, upgrade.output);
    assert.doesNotMatch(upgrade.output, /EBUSY|EPERM/);
    assert.deepEqual(fs.readdirSync(scope).filter((name) => name.startsWith('.')), []);
    assert.equal(installed, versions[1]);
    assert.equal(server.exitCode, null, 'running server is unaffected by the upgrade');

    server.stdin.end();
    await waitForExit(server);
  },
);
