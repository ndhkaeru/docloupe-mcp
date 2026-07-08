#!/usr/bin/env node
'use strict';

const fs = require('fs');
const https = require('https');
const os = require('os');
const path = require('path');

const OWNER = process.env.GITHUB_REPOSITORY_OWNER || 'ndhkaeru';
const REPO = (process.env.GITHUB_REPOSITORY || 'ndhkaeru/docforge-mcp').split('/')[1] || 'docforge-mcp';
const TAG = process.env.GITHUB_REF_NAME || process.argv[2] || 'latest';
const PACKAGE_ROOT = path.resolve(__dirname, '..');
const NATIVE_DIR = path.join(PACKAGE_ROOT, 'native');

const TARGETS = [
  ['win32-x64', 'windows-x64', '.exe'],
  ['linux-x64', 'linux-x64', ''],
  ['darwin-x64', 'macos-x64', ''],
  ['darwin-arm64', 'macos-arm64', ''],
];
const SERVERS = ['excel', 'md', 'pdf', 'docx', 'pptx', 'csv', 'html', 'text', 'json'];

function requestJson(url) {
  return new Promise((resolve, reject) => {
    const headers = {
      'User-Agent': 'docforge-mcp-release-packager',
      'Accept': 'application/vnd.github+json',
    };
    if (process.env.GITHUB_TOKEN) headers.Authorization = `Bearer ${process.env.GITHUB_TOKEN}`;
    https.get(url, { headers }, (response) => {
      if (response.statusCode < 200 || response.statusCode >= 300) {
        reject(new Error(`GET ${url} failed: ${response.statusCode}`));
        response.resume();
        return;
      }
      let body = '';
      response.setEncoding('utf8');
      response.on('data', (chunk) => body += chunk);
      response.on('end', () => resolve(JSON.parse(body)));
    }).on('error', reject);
  });
}

function download(url, outputPath) {
  return new Promise((resolve, reject) => {
    const headers = { 'User-Agent': 'docforge-mcp-release-packager' };
    if (process.env.GITHUB_TOKEN) headers.Authorization = `Bearer ${process.env.GITHUB_TOKEN}`;
    https.get(url, { headers }, (response) => {
      if ([301, 302, 303, 307, 308].includes(response.statusCode)) {
        download(response.headers.location, outputPath).then(resolve, reject);
        return;
      }
      if (response.statusCode < 200 || response.statusCode >= 300) {
        reject(new Error(`download failed: ${response.statusCode} ${url}`));
        response.resume();
        return;
      }
      const file = fs.createWriteStream(outputPath);
      response.pipe(file);
      file.on('finish', () => file.close(resolve));
      file.on('error', reject);
    }).on('error', reject);
  });
}

async function main() {
  fs.rmSync(NATIVE_DIR, { recursive: true, force: true });
  fs.mkdirSync(NATIVE_DIR, { recursive: true });

  const releaseUrl = TAG === 'latest'
    ? `https://api.github.com/repos/${OWNER}/${REPO}/releases/latest`
    : `https://api.github.com/repos/${OWNER}/${REPO}/releases/tags/${TAG}`;
  const release = await requestJson(releaseUrl);

  for (const [npmPlatform, releasePlatform, suffix] of TARGETS) {
    const outDir = path.join(NATIVE_DIR, npmPlatform);
    fs.mkdirSync(outDir, { recursive: true });
    for (const server of SERVERS) {
      const binaryName = `${server}-tools${suffix}`;
      const assetName = `docforge-mcp-${server}-tools-${releasePlatform}${suffix}`;
      const asset = release.assets.find((item) => item.name === assetName);
      if (!asset) throw new Error(`Missing release asset: ${assetName}`);
      const outPath = path.join(outDir, binaryName);
      await download(asset.browser_download_url, outPath);
      if (suffix === '') fs.chmodSync(outPath, 0o755);
      console.log(`prepared ${npmPlatform}: ${binaryName}`);
    }
  }
}

main().catch((error) => {
  console.error(error.stack || error.message);
  process.exit(1);
});
