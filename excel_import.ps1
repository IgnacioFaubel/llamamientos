[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$WorkbookPath,
    [Parameter(Mandatory = $true)][string]$SheetName,
    [Parameter(Mandatory = $true)][string]$TableName,
    [Parameter(Mandatory = $true)][string]$DataPath
)

$ErrorActionPreference = "Stop"
$excel = $null
$workbook = $null
$worksheet = $null
$table = $null

try {
    $payload = Get-Content -LiteralPath $DataPath -Raw -Encoding UTF8 | ConvertFrom-Json
    $records = @($payload.records)
    if ($records.Count -lt 1) {
        throw "No hay registros para importar."
    }

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AskToUpdateLinks = $false
    $excel.AutomationSecurity = 3

    $expectedPath = [System.IO.Path]::GetFullPath($WorkbookPath)
    $workbook = $excel.Workbooks.Open($expectedPath, 0, $false)
    $openedPath = [System.IO.Path]::GetFullPath([string]$workbook.FullName)
    if (-not [string]::Equals($openedPath, $expectedPath, [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "Excel no ha abierto el archivo original seleccionado. Se cancela la importacion."
    }
    if ($workbook.ReadOnly) {
        throw "El archivo original esta abierto como solo lectura o bloqueado. Se cancela la importacion."
    }
    $excel.Calculation = -4135
    $worksheet = $workbook.Worksheets.Item($SheetName)
    $table = $worksheet.ListObjects.Item($TableName)

    if ($null -eq $table.DataBodyRange) {
        throw "La tabla $TableName no tiene una fila modelo con formulas."
    }

    $formulaTemplates = @()
    for ($column = 6; $column -le 11; $column++) {
        $formula = [string]$table.DataBodyRange.Cells.Item(1, $column).FormulaR1C1
        if ([string]::IsNullOrWhiteSpace($formula) -or -not $formula.StartsWith("=")) {
            throw "Falta la formula modelo de la columna $column en la tabla $TableName."
        }
        $formulaTemplates += $formula
    }

    $headerRow = $table.HeaderRowRange.Row
    $firstColumn = $table.Range.Column
    $oldBody = $table.DataBodyRange
    $oldBody.ClearContents()

    $newRange = $worksheet.Range(
        $worksheet.Cells.Item($headerRow, $firstColumn),
        $worksheet.Cells.Item($headerRow + $records.Count, $firstColumn + 10)
    )
    $table.Resize($newRange)

    $values = [System.Array]::CreateInstance([object], @($records.Count, 5))
    $culture = [System.Globalization.CultureInfo]::InvariantCulture
    for ($row = 0; $row -lt $records.Count; $row++) {
        $record = @($records[$row])
        $workDate = [datetime]::ParseExact([string]$record[0], "yyyy-MM-dd", $culture)
        $values.SetValue($workDate.ToOADate(), $row, 0)
        for ($column = 1; $column -lt 5; $column++) {
            $values.SetValue($record[$column], $row, $column)
        }
    }

    $body = $table.DataBodyRange
    $dataRange = $worksheet.Range(
        $body.Cells.Item(1, 1),
        $body.Cells.Item($records.Count, 5)
    )
    $dataRange.Value2 = $values

    for ($index = 0; $index -lt 6; $index++) {
        $column = 6 + $index
        $formulaRange = $worksheet.Range(
            $body.Cells.Item(1, $column),
            $body.Cells.Item($records.Count, $column)
        )
        $formulaRange.FormulaR1C1 = $formulaTemplates[$index]
    }

    $excel.Calculation = -4105
    $excel.CalculateFullRebuild()
    $workbook.Save()
    if (-not $workbook.Saved) {
        throw "Excel no ha confirmado el guardado del archivo original."
    }
    $savedPath = [System.IO.Path]::GetFullPath([string]$workbook.FullName)
    if (-not [string]::Equals($savedPath, $expectedPath, [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "Excel ha intentado guardar en una ruta distinta del archivo original."
    }
    Write-Output "OK"
}
catch {
    [Console]::Error.WriteLine($_.Exception.Message)
    exit 1
}
finally {
    if ($null -ne $workbook) {
        try { $workbook.Close($false) } catch {}
    }
    if ($null -ne $excel) {
        try { $excel.Quit() } catch {}
    }
    foreach ($item in @($table, $worksheet, $workbook, $excel)) {
        if ($null -ne $item) {
            try { [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($item) } catch {}
        }
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
