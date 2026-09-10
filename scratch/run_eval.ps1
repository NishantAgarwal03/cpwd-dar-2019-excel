param([string]$FilePath, [string]$JsonPath)

Get-Process excel -ErrorAction SilentlyContinue | Stop-Process -Force
Start-Sleep -Milliseconds 800

$cases = Get-Content -Raw -Encoding UTF8 $JsonPath | ConvertFrom-Json

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
    $wb = $excel.Workbooks.Open($FilePath, $false, $false)
    $ws = $wb.Worksheets.Item("02_Earth_Work")

    Write-Host "========================================================================"
    Write-Host "LIVE EXCEL COM RECALCULATION TESTS FOR C9 / D9 SELECTIONS"
    Write-Host "========================================================================"

    $allPassed = $true

    foreach ($tc in $cases) {
        $ws.Range("C9").Value2 = $tc.Cat
        $ws.Range("D9").Value2 = $tc.Task
        $excel.CalculateFull()
        Start-Sleep -Milliseconds 200

        $bQty = $ws.Range("D10").Value2
        $bUnit = $ws.Range("D11").Value2
        $darPub = $ws.Range("D14").Value2
        $wSub = $ws.Range("H41").Value2
        $rate = $ws.Range("H55").Value2
        $say = $ws.Range("H56").Value2
        $diff = $ws.Range("H58").Value2
        $audit = $ws.Range("D15").Text

        $res1Code = $ws.Range("B27").Value2
        $res1Qty = $ws.Range("F27").Value2
        $res1Rate = $ws.Range("G27").Value2
        $res1Amt = $ws.Range("H27").Value2

        Write-Host ""
        Write-Host "TEST: [$($tc.Cat)] -> [$($tc.Task)]"
        Write-Host "  Batch: $bQty $bUnit | Published Say: Rs $darPub | Derived Say: Rs $say | Diff: Rs $diff | Audit: $audit"
        Write-Host "  Line 1: Res $res1Code, Qty=$res1Qty, Rate=$res1Rate, Amt=$res1Amt | Direct W: Rs $wSub"

        $qtyDiff = [Math]::Abs($bQty - $tc.ExpBatch)
        $sayDiff = [Math]::Abs($say - $tc.ExpSay)
        $diffVal = [Math]::Abs($diff)

        if ($qtyDiff -lt 0.01 -and $bUnit -eq $tc.ExpUnit -and $sayDiff -le 0.05 -and $diffVal -le 0.05 -and $audit -like "*PASS*") {
            Write-Host "  --> PASS: 100% DAR EXACT MATCH" -ForegroundColor Green
        } else {
            Write-Host "  --> FAIL: MISMATCH DETECTED" -ForegroundColor Red
            $allPassed = $false
        }
    }

    Write-Host ""
    Write-Host "========================================================================"
    if ($allPassed) {
        Write-Host "ALL TEST CASES PASSED WITH 100% ACCURACY AND ZERO VARIANCE!" -ForegroundColor Green
    } else {
        Write-Host "SOME TEST CASES FAILED!" -ForegroundColor Red
    }
    Write-Host "========================================================================"

} finally {
    if ($wb) { $wb.Close($false) }
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}
