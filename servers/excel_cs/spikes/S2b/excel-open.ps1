param(
    [string]$SourceRoot = 'D:\data-test\excel-preservation-fixtures\sources',
    [string]$Suffix = ''
)

$ErrorActionPreference = 'Stop'
$files = @(
    '02-table-metadata-source.xlsx',
    '05-advanced-package-source.xlsm',
    '07-real-package-source.xlsx'
)
$excel = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AskToUpdateLinks = $false
    $excel.EnableEvents = $false
    $excel.AutomationSecurity = 3
    foreach ($baseName in $files) {
        $name = [IO.Path]::GetFileNameWithoutExtension($baseName) + $Suffix + [IO.Path]::GetExtension($baseName)
        $path = Join-Path $SourceRoot $name
        $before = (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash
        $workbook = $null
        try {
            $workbook = $excel.Workbooks.Open($path, 0, $true)
            $readOnly = [bool]$workbook.ReadOnly
            $workbook.Close($false)
            Write-Output "$name | open=YES | read_only=$readOnly | source_hash_same=$($before -eq (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash)"
        }
        catch {
            Write-Output "$name | open=NO | hresult=$('{0:X8}' -f $_.Exception.HResult) | source_hash_same=$($before -eq (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash)"
        }
        finally {
            if ($null -ne $workbook) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($workbook) }
        }
    }
}
finally {
    if ($null -ne $excel) {
        $excel.Quit()
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)
    }
}
