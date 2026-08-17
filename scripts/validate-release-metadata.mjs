import fs from 'node:fs';
import path from 'node:path';
import process from 'node:process';
import { fileURLToPath } from 'node:url';

const scriptDirectory = path.dirname(fileURLToPath(import.meta.url));
const repository = path.resolve(scriptDirectory, '..');
const packageMetadata = JSON.parse(
  fs.readFileSync(path.join(repository, 'packages', 'npm', 'package.json'), 'utf8'),
);
const serverMetadata = JSON.parse(
  fs.readFileSync(path.join(repository, 'server.json'), 'utf8'),
);
const releaseTag = process.argv[2] || process.env.RELEASE_TAG;
const mismatches = [];

if (!releaseTag) {
  mismatches.push('release tag is required');
} else if (!/^v\d+\.\d+\.\d+(?:-[0-9A-Za-z.-]+)?$/.test(releaseTag)) {
  mismatches.push(`release tag ${releaseTag} is not a supported semantic version tag`);
}

const expectedTag = `v${packageMetadata.version}`;
if (releaseTag && releaseTag !== expectedTag) {
  mismatches.push(`release tag ${releaseTag} != package tag ${expectedTag}`);
}
if (serverMetadata.version !== packageMetadata.version) {
  mismatches.push(
    `server.version=${serverMetadata.version} != package.version=${packageMetadata.version}`,
  );
}
if (serverMetadata.packages?.[0]?.identifier !== packageMetadata.name) {
  mismatches.push(
    `server package identifier=${serverMetadata.packages?.[0]?.identifier} != package.name=${packageMetadata.name}`,
  );
}
if (serverMetadata.packages?.[0]?.version !== packageMetadata.version) {
  mismatches.push(
    `server package version=${serverMetadata.packages?.[0]?.version} != package.version=${packageMetadata.version}`,
  );
}
if (packageMetadata.mcpName !== serverMetadata.name) {
  mismatches.push(
    `package mcpName=${packageMetadata.mcpName} != server.name=${serverMetadata.name}`,
  );
}

if (mismatches.length) {
  console.error('Release metadata validation failed:');
  for (const mismatch of mismatches) console.error(`  ${mismatch}`);
  process.exit(1);
}

console.log(`Release metadata is valid for ${releaseTag}.`);
