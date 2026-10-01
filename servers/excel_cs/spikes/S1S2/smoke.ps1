param(
    [string]$Binary = (Join-Path $PSScriptRoot 'out\aot\S1S2.exe'),
    [string]$Fixture = 'D:\data-test\excel-preservation-fixtures\sources\00-base.xlsx',
    [int]$Runs = 10
)

$ErrorActionPreference = 'Stop'
$samples = @()
for ($index = 0; $index -lt $Runs; $index++) {
    $startInfo = [System.Diagnostics.ProcessStartInfo]::new($Binary)
    $startInfo.UseShellExecute = $false
    $startInfo.CreateNoWindow = $true
    $startInfo.RedirectStandardInput = $true
    $startInfo.RedirectStandardOutput = $true
    $startInfo.RedirectStandardError = $true
    $process = [System.Diagnostics.Process]::new()
    $process.StartInfo = $startInfo
    $timer = [System.Diagnostics.Stopwatch]::StartNew()
    try {
        if (-not $process.Start()) { throw 'Could not start MCP server' }
        $request = @{ jsonrpc = '2.0'; id = 1; method = 'initialize'; params = @{ protocolVersion = '2025-06-18'; capabilities = @{}; clientInfo = @{ name = 's1s2-smoke'; version = '0.1' } } } | ConvertTo-Json -Depth 10 -Compress
        $process.StandardInput.WriteLine($request)
        $answer = $process.StandardOutput.ReadLineAsync()
        if (-not $answer.Wait(15000)) { throw 'Initialize timed out' }
        if ($null -eq $answer.Result) {
            $process.WaitForExit(3000) | Out-Null
            throw "Initialize EOF: exit=$($process.ExitCode) stderr=$($process.StandardError.ReadToEnd())"
        }
        $initialize = $answer.Result | ConvertFrom-Json
        if (-not $initialize.result) { throw "Initialize failed: $($answer.Result)" }
        $initializedMs = $timer.Elapsed.TotalMilliseconds
        $process.StandardInput.WriteLine('{"jsonrpc":"2.0","method":"notifications/initialized"}')
        $call = @{ jsonrpc = '2.0'; id = 2; method = 'tools/call'; params = @{ name = 'read_first_sheet'; arguments = @{ path = $Fixture } } } | ConvertTo-Json -Depth 10 -Compress
        $process.StandardInput.WriteLine($call)
        $answer = $process.StandardOutput.ReadLineAsync()
        if (-not $answer.Wait(15000)) { throw 'Tool call timed out' }
        if ($null -eq $answer.Result) {
            $process.WaitForExit(3000) | Out-Null
            throw "Tool EOF: exit=$($process.ExitCode) stderr=$($process.StandardError.ReadToEnd())"
        }
        $result = $answer.Result | ConvertFrom-Json
        if ($result.error -or $result.result.isError -or -not $result.result.structuredContent) { throw "Tool failed or missing structuredContent: $($answer.Result)" }
        $sample = [pscustomobject]@{ initialize_ms = [Math]::Round($initializedMs, 1); tool_ms = [Math]::Round($timer.Elapsed.TotalMilliseconds, 1); structured = $result.result.structuredContent }
        $samples += $sample
        if ($index -eq 0) { $sample | ConvertTo-Json -Depth 10 -Compress }
    }
    finally {
        if (-not $process.HasExited) { $process.Kill(); $process.WaitForExit(5000) | Out-Null }
        $process.Dispose()
    }
}
$ordered = @($samples.initialize_ms | Sort-Object)
$summary = [pscustomobject]@{
    runs = $Runs
    initialize_min_ms = $ordered[0]
    initialize_median_ms = $ordered[[int][Math]::Floor(($Runs - 1) / 2)]
    initialize_max_ms = $ordered[-1]
    tool_median_ms = @($samples.tool_ms | Sort-Object)[[int][Math]::Floor(($Runs - 1) / 2)]
}
$summary | ConvertTo-Json -Compress
