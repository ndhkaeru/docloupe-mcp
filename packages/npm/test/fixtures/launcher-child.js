'use strict';

const fs = require('fs');
const path = require('path');
const { spawn } = require('child_process');

function ignore(signalName) {
  try {
    process.on(signalName, () => {});
  } catch {
    return;
  }
}

function writeJson(outputPath, value) {
  const temporaryPath = `${outputPath}.${process.pid}.tmp`;
  fs.writeFileSync(temporaryPath, JSON.stringify(value));
  fs.renameSync(temporaryPath, outputPath);
}

function runEcho(outputPath) {
  const chunks = [];
  process.stdin.on('data', (chunk) => chunks.push(Buffer.from(chunk)));
  process.stdin.on('end', () => {
    fs.writeFileSync(outputPath, Buffer.concat(chunks));
    process.exitCode = 7;
  });
  process.stdin.resume();
}

function runGrandchild() {
  ignore('SIGINT');
  ignore('SIGTERM');
  setInterval(() => {}, 1000);
}

function runTree(outputPath) {
  ignore('SIGINT');
  ignore('SIGTERM');
  const grandchild = spawn(process.execPath, [__filename, 'grandchild'], {
    stdio: 'ignore',
    windowsHide: true,
  });
  writeJson(outputPath, { parent: process.pid, grandchild: grandchild.pid });
  process.stdin.on('end', () => {});
  process.stdin.resume();
  setInterval(() => {}, 1000);
}

function runSignal(readyPath, signalPath) {
  const finish = (signalName) => {
    fs.writeFileSync(signalPath, signalName);
    process.exitCode = 0;
    process.stdin.pause();
  };
  process.on('SIGINT', () => finish('SIGINT'));
  process.on('SIGTERM', () => finish('SIGTERM'));
  fs.writeFileSync(readyPath, String(process.pid));
  process.stdin.resume();
}

const [mode, ...args] = process.argv.slice(2);
if (mode === 'echo') {
  runEcho(path.resolve(args[0]));
} else if (mode === 'tree') {
  runTree(path.resolve(args[0]));
} else if (mode === 'grandchild') {
  runGrandchild();
} else if (mode === 'signal') {
  runSignal(path.resolve(args[0]), path.resolve(args[1]));
} else {
  process.exitCode = 2;
}
