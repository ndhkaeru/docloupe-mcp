<#
.SYNOPSIS
    Drives a DocLoupe Excel MCP server binary over stdio and checks a full
    load -> edit -> save -> reload round trip.

.DESCRIPTION
    Written for Windows PowerShell 5.1 so it runs inside a pristine Windows Sandbox
    (no Python, Node or .NET runtime installed). It also runs on a normal Windows host,
    which only proves that the harness works, not that the binary is self-contained.

    <InDir> must contain the server binary and fixture.xlsx. The binary is copied to a
    local temp directory before it runs. Results go to <OutDir>:
      result.json        status + machine inventory + every step (rewritten after each step)
      summary.txt        human-readable summary
      server-stderr.log  everything the server wrote to stderr
      saved.xlsx         the workbook the server saved

    The script source is ASCII-only; Vietnamese literals are JSON escapes decoded at runtime.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$InDir,
    [Parameter(Mandatory = $true)][string]$OutDir,
    [string]$BinaryName = 'excel-tools.exe',
    [ValidateSet('legacy')][string]$Scenario = 'legacy',
    [int]$StartupTimeoutSeconds = 180,
    [int]$CallTimeoutSeconds = 300,
    [switch]$ShutdownWhenDone
)

Set-StrictMode -Version 2
$ErrorActionPreference = 'Stop'
$Utf8 = New-Object System.Text.UTF8Encoding($false)

# ---------------------------------------------------------------- helpers

function U([string]$Escaped) {
    # Decode a JSON-escaped string literal, e.g. U '\u0111\u00e3'.
    return ('"' + $Escaped + '"') | ConvertFrom-Json
}

function Limit([string]$Text, [int]$Max = 800) {
    if ($null -eq $Text) { return '' }
    if ($Text.Length -le $Max) { return $Text }
    return $Text.Substring(0, $Max) + '...'
}

function Show-CodePoints([string]$Text) {
    if ($null -eq $Text) { return '<null>' }
    return (($Text.ToCharArray() | ForEach-Object { 'U+{0:X4}' -f [int]$_ }) -join ' ')
}

$JsonSerializer = $null
try {
    Add-Type -AssemblyName System.Web.Extensions
    $JsonSerializer = New-Object System.Web.Script.Serialization.JavaScriptSerializer
    $JsonSerializer.MaxJsonLength = [int]::MaxValue
    $JsonSerializer.RecursionLimit = 1000
} catch { $JsonSerializer = $null }

function ConvertFrom-JsonExact([string]$Json) {
    # Keys are case-sensitive here. The legacy server emits keys that differ only by case
    # (e.g. 'vAlign' and 'valign'), which Windows PowerShell's ConvertFrom-Json rejects.
    if ($null -ne $script:JsonSerializer) { return , $script:JsonSerializer.DeserializeObject($Json) }
    return ($Json | ConvertFrom-Json -AsHashtable)
}

function Get-CommandPath([string]$Name) {
    $found = Get-Command $Name -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($null -eq $found) { return $null }
    return $found.Source
}

# ---------------------------------------------------------------- result bookkeeping

New-Item -ItemType Directory -Path $OutDir -Force | Out-Null
$ResultPath = Join-Path $OutDir 'result.json'
$Steps = New-Object System.Collections.ArrayList
$StartedAt = (Get-Date).ToUniversalTime().ToString('o')
$os = Get-CimInstance Win32_OperatingSystem
$Machine = [ordered]@{
    user          = $env:USERNAME
    in_sandbox    = ($env:USERNAME -eq 'WDAGUtilityAccount')
    os            = "$($os.Caption) $($os.Version)"
    powershell    = $PSVersionTable.PSVersion.ToString()
    python        = Get-CommandPath 'python'
    py            = Get-CommandPath 'py'
    node          = Get-CommandPath 'node'
    dotnet        = Get-CommandPath 'dotnet'
    dotnet_dir    = (Test-Path (Join-Path $env:ProgramFiles 'dotnet'))
    vc_redist_x64 = (Test-Path 'HKLM:\SOFTWARE\Microsoft\VisualStudio\14.0\VC\Runtimes\x64')
}

function Save-Result([string]$Status) {
    $document = [ordered]@{
        status     = $Status
        scenario   = $Scenario
        binary     = $BinaryName
        started_at = $StartedAt
        machine    = $Machine
        steps      = @($Steps)
    }
    [System.IO.File]::WriteAllText($ResultPath, ($document | ConvertTo-Json -Depth 10), $Utf8)
}

function Invoke-Step([string]$StepName, [scriptblock]$StepBody) {
    $stepWatch = [System.Diagnostics.Stopwatch]::StartNew()
    $stepOk = $true
    try { $stepValue = & $StepBody } catch { $stepOk = $false; $stepValue = $_.Exception.Message }
    $stepWatch.Stop()
    if ($stepValue -is [string]) { $shown = Limit $stepValue } else { $shown = Limit ($stepValue | ConvertTo-Json -Compress -Depth 5) }
    [void]$Steps.Add([ordered]@{ name = $StepName; ok = $stepOk; ms = $stepWatch.ElapsedMilliseconds; detail = $shown })
    Save-Result 'running'
    if ($stepOk) { $label = 'PASS' } else { $label = 'FAIL' }
    Write-Host ('[{0}] {1} ({2} ms) {3}' -f $label, $StepName, $stepWatch.ElapsedMilliseconds, (Limit $shown 200))
    if (-not $stepOk) { throw "Step '$StepName' failed: $stepValue" }
    return $stepValue
}

# ---------------------------------------------------------------- MCP stdio client

$Proc = $null
$PendingRead = $null
$StderrTask = $null
$NextId = 1

function Start-Server([string]$Exe, [string]$WorkDir) {
    # .NET Framework creates the child's stdin writer with Console.InputEncoding and AutoFlush,
    # which writes that encoding's preamble immediately. On a UTF-8 console that is a BOM in
    # front of the first JSON-RPC message, which the server rejects. Use a BOM-less encoding.
    try {
        if ([Console]::InputEncoding.GetPreamble().Length -gt 0) { [Console]::InputEncoding = $Utf8 }
    } catch { }
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = $Exe
    $psi.WorkingDirectory = $WorkDir
    $psi.UseShellExecute = $false
    $psi.CreateNoWindow = $true
    $psi.RedirectStandardInput = $true
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError = $true
    $psi.StandardOutputEncoding = $Utf8
    $psi.StandardErrorEncoding = $Utf8
    $script:Proc = [System.Diagnostics.Process]::Start($psi)
    # Drain stderr continuously so the server can never block on a full pipe.
    $script:StderrTask = $script:Proc.StandardError.ReadToEndAsync()
}

function Send-Message($Message) {
    $json = $Message | ConvertTo-Json -Depth 20 -Compress
    $bytes = $Utf8.GetBytes($json + "`n")
    $stream = $script:Proc.StandardInput.BaseStream
    $stream.Write($bytes, 0, $bytes.Length)
    $stream.Flush()
}

function Read-Response([int]$Id, [int]$TimeoutSeconds) {
    $deadline = [DateTime]::UtcNow.AddSeconds($TimeoutSeconds)
    while ($true) {
        if ($null -eq $script:PendingRead) { $script:PendingRead = $script:Proc.StandardOutput.ReadLineAsync() }
        $remaining = [int][Math]::Max(0, ($deadline - [DateTime]::UtcNow).TotalMilliseconds)
        if (-not $script:PendingRead.Wait($remaining)) {
            throw "Timed out after $TimeoutSeconds s waiting for response id $Id"
        }
        $line = $script:PendingRead.Result
        $script:PendingRead = $null
        if ($null -eq $line) {
            if ($script:Proc.HasExited) { $code = $script:Proc.ExitCode } else { $code = 'still running' }
            throw "Server closed stdout (exit code: $code)"
        }
        if ($line.Trim().Length -eq 0) { continue }
        try { $message = $line | ConvertFrom-Json } catch { throw "Non-JSON line on stdout: $(Limit $line 300)" }
        if ($message.PSObject.Properties['id'] -and $null -ne $message.id -and [int]$message.id -eq $Id) {
            return $message
        }
        # Notifications and unrelated messages are ignored.
    }
}

function Invoke-Rpc([string]$Method, $Params, [int]$TimeoutSeconds) {
    $id = $script:NextId
    $script:NextId++
    Send-Message ([ordered]@{ jsonrpc = '2.0'; id = $id; method = $Method; params = $Params })
    $response = Read-Response $id $TimeoutSeconds
    if ($response.PSObject.Properties['error']) {
        throw "JSON-RPC error $($response.error.code): $($response.error.message)"
    }
    return $response.result
}

function Invoke-Tool([string]$ToolName, $Arguments) {
    $result = Invoke-Rpc 'tools/call' ([ordered]@{ name = $ToolName; arguments = $Arguments }) $CallTimeoutSeconds
    $text = (@($result.content) | Where-Object { $_.type -eq 'text' } | ForEach-Object { $_.text }) -join "`n"
    if ($result.PSObject.Properties['isError'] -and $result.isError) {
        throw "Tool $ToolName returned an error: $(Limit $text 600)"
    }
    return $text
}

function Stop-Server {
    if ($null -eq $script:Proc) { return }
    try { $script:Proc.StandardInput.Close() } catch { }
    if (-not $script:Proc.WaitForExit(15000)) {
        # PyInstaller one-file builds run a child process; kill the whole tree.
        & taskkill.exe /T /F /PID $script:Proc.Id 2>&1 | Out-Null
    }
    $stderr = ''
    if ($null -ne $script:StderrTask -and $script:StderrTask.Wait(10000)) { $stderr = $script:StderrTask.Result }
    [System.IO.File]::WriteAllText((Join-Path $OutDir 'server-stderr.log'), $stderr, $Utf8)
}

function Get-SessionKey([string]$LoadText) {
    # Legacy output: Loaded: session_key='C:\\path\\file.xlsx' | sheets=[...] (Python repr).
    if ($LoadText -notmatch "session_key=(['""])(.*?)\1 \|") { throw "No session_key in: $(Limit $LoadText 300)" }
    return ($Matches[2] -replace '\\\\', '\')
}

function Assert-TextList([string[]]$Actual, [string[]]$Expected, [string]$What) {
    if ($Actual.Count -ne $Expected.Count) { throw "$What count: expected $($Expected.Count), got $($Actual.Count)" }
    for ($i = 0; $i -lt $Expected.Count; $i++) {
        if ($Actual[$i] -cne $Expected[$i]) {
            throw "$What[$i]: expected [$(Show-CodePoints $Expected[$i])] got [$(Show-CodePoints $Actual[$i])]"
        }
    }
}

# ---------------------------------------------------------------- scenarios

function Invoke-LegacyScenario([string]$Fixture, [string]$Saved) {
    # Tool names and output formats of the Python server (servers/excel, 1.1.x).
    $required = @('excel_load', 'excel_get_cell', 'excel_edit_cells', 'excel_save', 'excel_close', 'excel_validate_workbook')
    $edited = U '\u0111\u00e3 s\u1eeda - Ti\u1ebfng Vi\u1ec7t'
    $expectedRuns = @((U '\u0110\u1ecf \u0111\u1eadm '), (U 'nghi\u00eang '), (U 'g\u1ea1ch ch\u00e2n'))

    Invoke-Step 'tools/list contains required tools' {
        $listed = Invoke-Rpc 'tools/list' @{} 60
        $names = @($listed.tools | ForEach-Object { $_.name })
        $missing = @($required | Where-Object { $names -notcontains $_ })
        if ($missing.Count -gt 0) { throw "Missing tools: $($missing -join ', ')" }
        "$($names.Count) tools listed"
    } | Out-Null

    $key = Invoke-Step 'excel_load fixture' { Get-SessionKey (Invoke-Tool 'excel_load' @{ uri = $Fixture }) }

    Invoke-Step 'A1 rich text has 3 runs with exact Vietnamese text' {
        $cell = ConvertFrom-JsonExact (Invoke-Tool 'excel_get_cell' ([ordered]@{ session_key = $key; sheet_name = 'S'; row_index = 0; col_index = 0; include_rich_text = $true }))
        $texts = [string[]]@($cell.rich_text.runs | ForEach-Object { $_.text })
        Assert-TextList $texts $expectedRuns 'A1 run'
        "$($texts.Count) runs"
    } | Out-Null

    Invoke-Step 'excel_edit_cells B2 (Vietnamese text)' {
        # The documented A1 form first; older builds may only accept the 0-based forms.
        # Any rejected form is reported in the step detail, never hidden.
        $forms = @(
            @{ form = 'A1'; edits = @([ordered]@{ cell = 'B2'; value = $edited }) },
            @{ form = 'flat'; edits = @([ordered]@{ row_index = 1; col_index = 1; value = $edited }) },
            @{ form = 'grouped'; edits = @([ordered]@{ row_index = 1; edits = [ordered]@{ '1' = $edited } }) }
        )
        $rejected = @()
        foreach ($candidate in $forms) {
            try {
                Invoke-Tool 'excel_edit_cells' ([ordered]@{ session_key = $key; sheet_name = 'S'; edits = $candidate.edits }) | Out-Null
                if ($rejected.Count -eq 0) { return "edited using the $($candidate.form) form" }
                return "edited using the $($candidate.form) form; REJECTED: $($rejected -join ' | ')"
            } catch { $rejected += "$($candidate.form): $(Limit $_.Exception.Message 160)" }
        }
        throw "Every edit form was rejected: $($rejected -join ' | ')"
    } | Out-Null

    Invoke-Step 'excel_save with verify_preservation=true' {
        $report = ConvertFrom-JsonExact (Invoke-Tool 'excel_save' ([ordered]@{ session_key = $key; output_path = $Saved; report_format = 'json'; verify_preservation = $true }))
        if (-not (Test-Path -LiteralPath $Saved)) { throw 'The saved workbook does not exist.' }
        $verification = $report.verification
        if ($verification.status -ne 'completed' -or -not $verification.preservation_ok) {
            throw "Verification not clean: $(Limit ($verification | ConvertTo-Json -Compress -Depth 4) 600)"
        }
        "verification=$($verification.status) preservation_ok=$($verification.preservation_ok) size=$((Get-Item -LiteralPath $Saved).Length)"
    } | Out-Null

    Invoke-Step 'excel_close' { Invoke-Tool 'excel_close' @{ session_key = $key } | Out-Null; 'closed' } | Out-Null

    $key2 = Invoke-Step 'excel_load saved workbook' { Get-SessionKey (Invoke-Tool 'excel_load' @{ uri = $Saved }) }

    Invoke-Step 'B2 reads back exactly' {
        $cell = ConvertFrom-JsonExact (Invoke-Tool 'excel_get_cell' ([ordered]@{ session_key = $key2; sheet_name = 'S'; row_index = 1; col_index = 1 }))
        Assert-TextList ([string[]]@([string]$cell.value)) ([string[]]@($edited)) 'B2'
        'B2 matches'
    } | Out-Null

    Invoke-Step 'A1 rich text still has 3 runs after save' {
        $cell = ConvertFrom-JsonExact (Invoke-Tool 'excel_get_cell' ([ordered]@{ session_key = $key2; sheet_name = 'S'; row_index = 0; col_index = 0; include_rich_text = $true }))
        $texts = [string[]]@($cell.rich_text.runs | ForEach-Object { $_.text })
        Assert-TextList $texts $expectedRuns 'A1 run'
        "$($texts.Count) runs"
    } | Out-Null

    Invoke-Step 'excel_validate_workbook saved workbook' {
        $report = ConvertFrom-JsonExact (Invoke-Tool 'excel_validate_workbook' @{ path = $Saved })
        if (-not $report.valid) { throw "Package invalid: $(Limit ($report | ConvertTo-Json -Compress -Depth 4) 600)" }
        'valid'
    } | Out-Null

    Invoke-Step 'excel_close saved session' { Invoke-Tool 'excel_close' @{ session_key = $key2 } | Out-Null; 'closed' } | Out-Null
}

# ---------------------------------------------------------------- main

$status = 'failed'
$workDir = Join-Path $env:TEMP ('docloupe-xt-' + [Guid]::NewGuid().ToString('N').Substring(0, 8))
$saved = Join-Path $workDir 'saved.xlsx'
try {
    $exe = Invoke-Step 'copy binary and fixture to a local directory' {
        New-Item -ItemType Directory -Path $workDir -Force | Out-Null
        Copy-Item -LiteralPath (Join-Path $InDir $BinaryName) -Destination $workDir
        Copy-Item -LiteralPath (Join-Path $InDir 'fixture.xlsx') -Destination $workDir
        Join-Path $workDir $BinaryName
    }

    Invoke-Step 'start server and initialize' {
        Start-Server $exe $workDir
        $init = Invoke-Rpc 'initialize' ([ordered]@{
                protocolVersion = '2025-06-18'
                capabilities    = @{}
                clientInfo      = [ordered]@{ name = 'docloupe-clean-machine-test'; version = '1' }
            }) $StartupTimeoutSeconds
        Send-Message ([ordered]@{ jsonrpc = '2.0'; method = 'notifications/initialized' })
        "server=$($init.serverInfo.name) $($init.serverInfo.version) protocol=$($init.protocolVersion)"
    } | Out-Null

    switch ($Scenario) {
        'legacy' { Invoke-LegacyScenario (Join-Path $workDir 'fixture.xlsx') $saved }
    }
    $status = 'passed'
} catch {
    Write-Host "Stopped: $($_.Exception.Message)"
} finally {
    Stop-Server
    if (Test-Path -LiteralPath $saved) { Copy-Item -LiteralPath $saved -Destination (Join-Path $OutDir 'saved.xlsx') -Force }
    # The work directory only holds a copy of the binary and scratch files; artifacts are in OutDir.
    try { Remove-Item -LiteralPath $workDir -Recurse -Force -ErrorAction Stop } catch { }
    Save-Result $status
    $lines = @("status: $status", "machine: $($Machine | ConvertTo-Json -Compress)")
    foreach ($step in $Steps) {
        if ($step.ok) { $mark = 'PASS' } else { $mark = 'FAIL' }
        $lines += ('{0}  {1,7} ms  {2}  {3}' -f $mark, $step.ms, $step.name, (Limit $step.detail 300))
    }
    [System.IO.File]::WriteAllText((Join-Path $OutDir 'summary.txt'), ($lines -join "`r`n"), $Utf8)
}

# Only ever shut down inside Windows Sandbox, and only when explicitly asked.
if ($ShutdownWhenDone -and $env:USERNAME -eq 'WDAGUtilityAccount') {
    & shutdown.exe /s /t 5 | Out-Null
}

if ($status -eq 'passed') { exit 0 } else { exit 1 }
