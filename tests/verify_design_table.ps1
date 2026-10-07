param([Parameter(Mandatory = $true)][string]$WorkbookPath)

$ErrorActionPreference = 'Stop'
$excel = $null
$book = $null
$sheet = $null

function Assert-Equal($Actual, $Expected, [string]$Message) {
    if ($Actual -ne $Expected) {
        throw "$Message`: expected [$Expected], got [$Actual]"
    }
}

function Read-Number([string]$Address) {
    $text = [string]$sheet.Range($Address).Value2
    return [double]::Parse($text.TrimStart('='), [Globalization.CultureInfo]::InvariantCulture)
}

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AutomationSecurity = 3
    $book = $excel.Workbooks.Open($WorkbookPath, 0, $true)
    $sheet = $book.Worksheets.Item(1)
    $originals = $sheet.Range('N3:N9').Formula
    $format = $sheet.Range('N8').NumberFormat
    $allFormulas = $sheet.UsedRange.Formula
    $sheet.Range('N8').NumberFormat = '@'

    $wallLabels = @('ONE WALL', 'TWO WALLS (CENTER AND LEFT SIDE)',
        'TWO WALLS (CENTER AND RIGHT SIDE)', 'THREE WALLS', 'FOUR WALLS')
    $wallCodes = @(1, 2, 2, 3, 4)
    $sideCodes = @(0, 2, 1, 0, 0)
    $returns = @('Default', 'single_return1', 'single_return2', 'double_return')
    $gauges = @('3/16', '1/4')
    $thicknesses = @(0.1875, 0.25)
    $materials = @('MILD STEEL', 'BORCON WEATHERING STEEL')
    $count = 0

    foreach ($dimensions in @(@(60.0, 60.0, 10.0), @(500.0, 500.0, 48.0), @(121.0, 120.0, 6.0))) {
        for ($w = 0; $w -lt 5; $w++) {
            for ($r = 0; $r -lt 4; $r++) {
                for ($g = 0; $g -lt 2; $g++) {
                    for ($m = 0; $m -lt 2; $m++) {
                        $inputs = @($wallLabels[$w], $returns[$r], $dimensions[0],
                            $dimensions[1], $dimensions[2], $gauges[$g], $materials[$m])
                        $matrix = New-Object 'object[,]' 7, 1
                        for ($i = 0; $i -lt 7; $i++) {
                            $matrix[$i, 0] = $inputs[$i]
                        }
                        $sheet.Range('N3:N9').Value2 = $matrix
                        $sheet.Calculate()
                        Assert-Equal (Read-Number 'D3') $wallCodes[$w] 'Wall code'
                        Assert-Equal (Read-Number 'C3') $sideCodes[$w] 'Side code'
                        Assert-Equal (Read-Number 'H3') $r 'Return code'
                        Assert-Equal (Read-Number 'E3') $dimensions[0] 'Length'
                        Assert-Equal (Read-Number 'F3') $dimensions[1] 'Width'
                        Assert-Equal (Read-Number 'G3') $dimensions[2] 'Height'
                        Assert-Equal (Read-Number 'K3') $thicknesses[$g] 'Physical thickness'
                        Assert-Equal (Read-Number 'R8') ($g + 1) 'Naming thickness'
                        Assert-Equal (Read-Number 'L3') ($m + 1) 'Material code'
                        $centerTotal = $dimensions[0]
                        if ($wallCodes[$w] -gt 2) {
                            $centerTotal -= $dimensions[2]
                        } elseif ($wallCodes[$w] -eq 2) {
                            $centerTotal -= $dimensions[2] / 2
                        }
                        $center = $centerTotal / [Math]::Ceiling($centerTotal / 120)
                        $segments = [Math]::Ceiling($dimensions[1] / 120)
                        $side = $dimensions[1] / $segments
                        if ($segments -gt 1) {
                            $side += $thicknesses[$g] + $thicknesses[$g] * ($segments - 2) / $segments
                        }
                        if ([Math]::Abs((Read-Number 'I3') - $center) -gt 0.000001) {
                            throw 'Center formula result differs from workbook rule'
                        }
                        if ([Math]::Abs((Read-Number 'J3') - $side) -gt 0.000001) {
                            throw 'Side formula result differs from workbook rule'
                        }
                        $count++
                    }
                }
            }
        }
    }
    Assert-Equal $count 240 'Scenario count'

    $sheet.Range('N3:N9').Formula = $originals
    $sheet.Range('N8').NumberFormat = $format
    $sheet.Calculate()
    $restored = $sheet.UsedRange.Formula
    for ($row = 1; $row -le $allFormulas.GetLength(0); $row++) {
        for ($column = 1; $column -le $allFormulas.GetLength(1); $column++) {
            Assert-Equal $restored[$row, $column] $allFormulas[$row, $column] "Formula/input at $row,$column"
        }
    }
    Assert-Equal (Read-Number 'I3') 113 'Original center result'
    Assert-Equal (Read-Number 'J3') 100.4 'Original side result'
    Write-Output 'PASS: 240 Excel formula scenarios; all original inputs and formulas restored; workbook not saved.'
} finally {
    if ($null -ne $book) { $book.Close($false) }
    if ($null -ne $sheet) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($sheet) }
    if ($null -ne $book) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($book) }
    if ($null -ne $excel) {
        $excel.Quit()
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)
    }
}
