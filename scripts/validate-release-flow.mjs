import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const repositoryRoot = join(dirname(fileURLToPath(import.meta.url)), '..');
const npmWorkflowPath = join(repositoryRoot, '.github', 'workflows', 'npm.yml');
const releaseWorkflowPath = join(repositoryRoot, '.github', 'workflows', 'release.yml');
const pushWorkflowPath = join(repositoryRoot, '.github', 'workflows', 'push.yml');
const npmWorkflow = readFileSync(npmWorkflowPath, 'utf8').replace(/\r\n?/g, '\n');
const releaseWorkflow = readFileSync(releaseWorkflowPath, 'utf8').replace(/\r\n?/g, '\n');
const pushWorkflow = readFileSync(pushWorkflowPath, 'utf8').replace(/\r\n?/g, '\n');
const failures = [];

function requireMatch(label, content, pattern) {
  if (!pattern.test(content)) failures.push(label);
}

function forbidMatch(label, content, pattern) {
  if (pattern.test(content)) failures.push(label);
}

requireMatch(
  'npm.yml must expose workflow_call for deterministic release chaining',
  npmWorkflow,
  /^  workflow_call:\s*$/m,
);
requireMatch(
  'npm.yml must keep workflow_dispatch for manual retries',
  npmWorkflow,
  /^  workflow_dispatch:\s*$/m,
);
forbidMatch(
  'npm.yml must not rely on release: published events created with GITHUB_TOKEN',
  npmWorkflow,
  /^  release:\s*$/m,
);
requireMatch(
  'npm.yml must derive RELEASE_TAG from reusable/manual inputs',
  npmWorkflow,
  /^      RELEASE_TAG: \$\{\{ inputs\.tag \}\}\s*$/m,
);
requireMatch(
  'npm.yml must fail when publishing is requested without NPM_TOKEN',
  npmWorkflow,
  /- name: Fail without npm token[\s\S]*?exit 1/,
);
requireMatch(
  'npm.yml must verify the package is publicly readable after publication',
  npmWorkflow,
  /- name: Verify public npm package[\s\S]*?npm pack "@ndhkaeru\/docloupe-mcp@\$\{version\}"/,
);

const publishJobStart = releaseWorkflow.indexOf('\n  publish:\n');
if (publishJobStart < 0) {
  failures.push('release.yml must define a publish job');
} else {
  const publishJob = releaseWorkflow.slice(publishJobStart);
  requireMatch(
    'release publish job must wait for the GitHub Release job',
    publishJob,
    /^    needs: release\s*$/m,
  );
  requireMatch(
    'release publish job must call the reusable npm workflow',
    publishJob,
    /^    uses: \.\/\.github\/workflows\/npm\.yml\s*$/m,
  );
  requireMatch(
    'release publish job must pass the release tag',
    publishJob,
    /^      tag: \$\{\{ github\.event_name == 'workflow_dispatch' && inputs\.tag \|\| github\.ref_name \}\}\s*$/m,
  );
  requireMatch(
    'release publish job must publish stable tags to npm',
    publishJob,
    /^      publish_npm: \$\{\{ !contains\(github\.event_name == 'workflow_dispatch' && inputs\.tag \|\| github\.ref_name, '-'\) \}\}\s*$/m,
  );
  requireMatch(
    'release publish job must inherit repository publication secrets',
    publishJob,
    /^    secrets: inherit\s*$/m,
  );
}

forbidMatch(
  'release.yml must not duplicate npm publish implementation',
  releaseWorkflow,
  /npm publish --access public/,
);

forbidMatch(
  'release.yml must not overwrite audited server sources with a broad overlay',
  releaseWorkflow,
  /Copy-Item[^\n]*source-overlay\/servers\/\*/,
);
requireMatch(
  'release.yml must restrict overlay copies to the converter allowlist',
  releaseWorkflow,
  /Copy-Item -LiteralPath \(Join-Path 'source-overlay' \$path\) -Destination \$path -Recurse -Force/,
);
requireMatch(
  'release.yml must package the Rust diagram renderer',
  releaseWorkflow,
  /--add-binary", "\$renderer\$\(\[IO\.Path\]::PathSeparator\)\."/,
);
requireMatch(
  'release.yml must smoke test packaged MCP servers',
  releaseWorkflow,
  /- name: Smoke test packaged MCP servers[\s\S]*?tests\/test_md_binary_smoke\.py tests\/test_excel_binary_smoke\.py/,
);
forbidMatch(
  'push.yml must not overwrite audited server sources with a broad overlay',
  pushWorkflow,
  /Copy-Item[^\n]*source-overlay\/servers\/\*/,
);
requireMatch(
  'push.yml must restrict overlay copies to the converter allowlist',
  pushWorkflow,
  /Copy-Item -LiteralPath \(Join-Path 'source-overlay' \$path\) -Destination \$path -Recurse -Force/,
);
requireMatch(
  'push.yml must build the Rust renderer before testing Markdown',
  pushWorkflow,
  /- name: Build portable diagram renderer[\s\S]*?cargo build --locked --release --manifest-path servers\/md\/diagram_renderer\/Cargo\.toml[\s\S]*?- name: Run tests/,
);

if (failures.length) {
  console.error('Release workflow validation failed:');
  for (const failure of failures) console.error(`  ${failure}`);
  process.exit(1);
}

console.log('Release workflow contract is valid.');
