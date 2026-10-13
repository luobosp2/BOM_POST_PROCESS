param([Parameter(Mandatory=$true)][string]$JobFile)
# BOM layout worker. ASCII source deliberately supports Windows PowerShell 5.1.
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = New-Object System.Text.UTF8Encoding($false)
$job = Get-Content -LiteralPath $JobFile -Raw -Encoding UTF8 | ConvertFrom-Json
$word = $null
$doc = $null
try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $word.AutomationSecurity = 3
    $doc = $word.Documents.Open([string]$job.input, $false, $true, $false)
    $doc.ActiveWindow.View.Type = 3
    $doc.Repaginate()
    if ($job.pdf) { $doc.ExportAsFixedFormat([string]$job.pdf, 17); $doc.Repaginate() }
    $result = [ordered]@{ pages=$doc.ComputeStatistics(2); tables=@() }
    foreach ($table in $doc.Tables) {
        $start = $table.Range.Duplicate
        $start.Collapse(1)
        $end = $doc.Range($table.Range.End-1, $table.Range.End-1)
        $result.tables += [ordered]@{
            start_page=$start.Information(3); end_page=$end.Information(3)
            start_y=$start.Information(6); end_y=$end.Information(6)
            rows=$table.Rows.Count
        }
    }
    if ($job.output) { $doc.SaveAs2([string]$job.output, 16) }
    $result | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $job.result -Encoding UTF8
} catch {
    [Console]::Error.WriteLine($_.Exception.Message)
    exit 1
} finally {
    if ($null -ne $doc) { $doc.Close([ref]0) }
    if ($null -ne $word) {
        $word.Quit([ref]0)
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($word)
    }
}
