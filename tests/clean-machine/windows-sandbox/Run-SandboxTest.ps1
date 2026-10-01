<#
.SYNOPSIS
    Runs the Excel MCP server binary inside Windows Sandbox (a pristine Windows with no
    Python, Node or .NET runtime) and reports whether a load -> edit -> save -> reload
    round trip works there.

.DESCRIPTION
    1. Stages the binary, the in-sandbox test script and a generated fixture in %TEMP%.
    2. Writes a .wsb configuration: networking disabled (offline machine), input folder
       mapped read-only, output folder mapped read-write, and a logon command that runs
       Test-McpServer.ps1 and then shuts the sandbox down.
    3. Starts Windows Sandbox, waits for result.json, and prints the summary.

    -RunOnHost runs the same test directly on this machine instead. That only checks the
    harness: this machine has runtimes installed, so it proves nothing about clean machines.

.EXAMPLE
    .\Run-SandboxTest.ps1
    .\Run-SandboxTest.ps1 -RunOnHost
    .\Run-SandboxTest.ps1 -Binary C:\path\excel-tools.exe -KeepOpen
#>
[CmdletBinding()]
param(
    [string]$Binary,
    [ValidateSet('legacy')][string]$Scenario = 'legacy',
    [int]$TimeoutMinutes = 20,
    [switch]$KeepOpen,
    [switch]$EnableNetworking,
    [switch]$PrepareOnly,
    [switch]$RunOnHost
)

Set-StrictMode -Version 2
$ErrorActionPreference = 'Stop'

if (-not $Binary) { $Binary = Join-Path $PSScriptRoot '..\..\..\dist\excel-tools.exe' }
$Binary = [System.IO.Path]::GetFullPath($Binary)
if (-not (Test-Path -LiteralPath $Binary)) { throw "Binary not found: $Binary" }
$binaryName = Split-Path -Leaf $Binary

function Show-Result([string]$OutDir) {
    $resultPath = Join-Path $OutDir 'result.json'
    if (-not (Test-Path -LiteralPath $resultPath)) {
        Write-Host 'No result.json was produced.' -ForegroundColor Red
        return $false
    }
    $result = Get-Content -LiteralPath $resultPath -Raw -Encoding UTF8 | ConvertFrom-Json
    Write-Host ''
    Write-Host "Status   : $($result.status)" -ForegroundColor $(if ($result.status -eq 'passed') { 'Green' } else { 'Red' })
    Write-Host "Machine  : $($result.machine.os) | user=$($result.machine.user) | in_sandbox=$($result.machine.in_sandbox)"
    Write-Host "Runtimes : python=$($result.machine.python) py=$($result.machine.py) node=$($result.machine.node) dotnet=$($result.machine.dotnet) dotnet_dir=$($result.machine.dotnet_dir) vc_redist_x64=$($result.machine.vc_redist_x64)"
    foreach ($step in $result.steps) {
        $mark = if ($step.ok) { 'PASS' } else { 'FAIL' }
        $color = if ($step.ok) { 'Gray' } else { 'Red' }
        Write-Host ('  {0}  {1,7} ms  {2}  {3}' -f $mark, $step.ms, $step.name, $step.detail) -ForegroundColor $color
    }
    Write-Host "Artifacts: $OutDir"
    return ($result.status -eq 'passed')
}

# ---------------------------------------------------------------- stage

$stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$stage = Join-Path $env:TEMP "docloupe-sandbox-$stamp"
$inDir = Join-Path $stage 'xt-in'
$outDir = Join-Path $stage 'xt-out'
New-Item -ItemType Directory -Path $inDir, $outDir -Force | Out-Null
Copy-Item -LiteralPath $Binary -Destination $inDir
Copy-Item -LiteralPath (Join-Path $PSScriptRoot 'Test-McpServer.ps1') -Destination $inDir
& (Join-Path $PSScriptRoot 'New-TestWorkbook.ps1') -Path (Join-Path $inDir 'fixture.xlsx') | Out-Null
Write-Host "Staged in $stage (binary: $Binary)"

# ---------------------------------------------------------------- host-only harness check

if ($RunOnHost) {
    Write-Host 'Running on this host (harness check only, NOT a clean-machine result)...' -ForegroundColor Yellow
    # Windows PowerShell 5.1, the same engine that runs inside the sandbox.
    & "$env:WINDIR\System32\WindowsPowerShell\v1.0\powershell.exe" -NoProfile -ExecutionPolicy Bypass `
        -File (Join-Path $inDir 'Test-McpServer.ps1') -InDir $inDir -OutDir $outDir -BinaryName $binaryName -Scenario $Scenario
    $passed = Show-Result $outDir
    if ($passed) { exit 0 } else { exit 1 }
}

# ---------------------------------------------------------------- sandbox

$sandboxExe = Join-Path $env:WINDIR 'System32\WindowsSandbox.exe'
if (-not (Test-Path -LiteralPath $sandboxExe)) {
    Write-Host 'Windows Sandbox is not enabled on this machine.' -ForegroundColor Red
    Write-Host 'Enable it once from an elevated PowerShell, then restart Windows:'
    Write-Host '  Enable-WindowsOptionalFeature -Online -FeatureName Containers-DisposableClientVM -All'
    Write-Host 'Requirements: Windows 10/11 Pro, Enterprise or Education; virtualization enabled.'
    Write-Host "The staged files are kept in $stage; rerun this script after enabling."
    exit 2
}
if (Get-Process -Name 'WindowsSandbox*' -ErrorAction SilentlyContinue) {
    Write-Host 'A Windows Sandbox instance is already running; only one can run at a time. Close it and retry.' -ForegroundColor Red
    exit 3
}

$sandboxHome = 'C:\Users\WDAGUtilityAccount\Desktop'
$innerArgs = "-InDir $sandboxHome\xt-in -OutDir $sandboxHome\xt-out -BinaryName $binaryName -Scenario $Scenario"
if (-not $KeepOpen) { $innerArgs += ' -ShutdownWhenDone' }
$networking = if ($EnableNetworking) { 'Enable' } else { 'Disable' }
$esc = { param($value) [System.Security.SecurityElement]::Escape($value) }

$wsb = @"
<Configuration>
  <Networking>$networking</Networking>
  <vGPU>Disable</vGPU>
  <ClipboardRedirection>Disable</ClipboardRedirection>
  <MappedFolders>
    <MappedFolder>
      <HostFolder>$(& $esc $inDir)</HostFolder>
      <ReadOnly>true</ReadOnly>
    </MappedFolder>
    <MappedFolder>
      <HostFolder>$(& $esc $outDir)</HostFolder>
      <ReadOnly>false</ReadOnly>
    </MappedFolder>
  </MappedFolders>
  <LogonCommand>
    <Command>powershell.exe -NoProfile -ExecutionPolicy Bypass -File $sandboxHome\xt-in\Test-McpServer.ps1 $innerArgs</Command>
  </LogonCommand>
</Configuration>
"@
$wsbPath = Join-Path $stage 'excel-tools-test.wsb'
[System.IO.File]::WriteAllText($wsbPath, $wsb, (New-Object System.Text.UTF8Encoding($false)))
Write-Host "Sandbox configuration: $wsbPath"

if ($PrepareOnly) {
    Write-Host 'PrepareOnly: open the .wsb file to start the test manually.'
    exit 0
}

Start-Process -FilePath $sandboxExe -ArgumentList "`"$wsbPath`""
Write-Host "Windows Sandbox started; waiting up to $TimeoutMinutes min for results..."

$resultPath = Join-Path $outDir 'result.json'
$deadline = (Get-Date).AddMinutes($TimeoutMinutes)
$finished = $false
while ((Get-Date) -lt $deadline) {
    Start-Sleep -Seconds 5
    if (Test-Path -LiteralPath $resultPath) {
        try {
            $state = (Get-Content -LiteralPath $resultPath -Raw -Encoding UTF8 | ConvertFrom-Json).status
            if ($state -ne 'running') { $finished = $true; break }
        } catch { }   # the file may be mid-write; retry on the next tick
    }
}
if (-not $finished) { Write-Host "Timed out after $TimeoutMinutes min." -ForegroundColor Red }
$passed = Show-Result $outDir
if ($passed) { exit 0 } else { exit 1 }
