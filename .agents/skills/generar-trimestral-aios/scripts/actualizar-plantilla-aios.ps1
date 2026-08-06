[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$PlantillaPath,

    [Parameter(Mandatory = $true)]
    [string]$FechaCorte,

    [string]$ExcelProcessIdFile,

    [switch]$SoloValidar
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$culture = [System.Globalization.CultureInfo]::InvariantCulture
$styles = [System.Globalization.DateTimeStyles]::None
$fecha = [DateTime]::MinValue
if (-not [DateTime]::TryParseExact($FechaCorte, 'yyyy-MM-dd', $culture, $styles, [ref]$fecha)) {
    throw "FechaCorte debe tener formato AAAA-MM-DD. Valor recibido: $FechaCorte"
}

$plantilla = [System.IO.Path]::GetFullPath($PlantillaPath)
if (-not (Test-Path -LiteralPath $plantilla -PathType Leaf)) {
    throw "No se encontró la plantilla: $plantilla"
}
if ([System.IO.Path]::GetExtension($plantilla) -ine '.xlsm') {
    throw "La plantilla debe ser un libro habilitado para macros (.xlsm): $plantilla"
}

if (-not ('AiosExcelNative' -as [type])) {
    Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;

public static class AiosExcelNative
{
    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);
}
'@
}

$excel = $null
$workbook = $null
$caratula = $null
$baseAnual = $null
$cuentas = $null
$cell = $null
$saveChanges = $false

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AskToUpdateLinks = $false
    $excel.EnableEvents = $true
    $excel.AutomationSecurity = if ($SoloValidar) { 3 } else { 1 }

    if (-not [string]::IsNullOrWhiteSpace($ExcelProcessIdFile)) {
        [uint32]$excelPid = 0
        [void][AiosExcelNative]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd, [ref]$excelPid)
        [System.IO.File]::WriteAllText([System.IO.Path]::GetFullPath($ExcelProcessIdFile), [string]$excelPid)
    }

    $workbook = $excel.Workbooks.Open($plantilla, 3, [bool]$SoloValidar)
    if (-not $SoloValidar -and $workbook.ReadOnly) {
        throw 'Excel abrió la plantilla como solo lectura. Cierre cualquier ventana que tenga abierto Plantilla AIOS-probable.xlsm y vuelva a intentar.'
    }

    $caratula = $workbook.Worksheets.Item('CARATULA')
    $macroAction = ''
    foreach ($shape in @($caratula.Shapes)) {
        try {
            if ([string]$shape.OnAction -match '(?i)(^|!)bajar$') {
                $macroAction = [string]$shape.OnAction
                break
            }
        } catch {
        }
    }
    if ([string]::IsNullOrWhiteSpace($macroAction)) {
        throw "No se encontró en CARATULA un botón asignado a la macro 'bajar'."
    }

    if ($SoloValidar) {
        Write-Output "VALIDACION_PLANTILLA_OK ruta=$plantilla macro=$macroAction fechaSolicitada=$FechaCorte"
        return
    }

    $caratula.Range('B2').Value2 = $fecha.ToOADate()
    $caratula.Range('B2').NumberFormat = 'dd/mm/yyyy'
    $excel.CalculateFull()

    $escapedWorkbookName = $workbook.Name.Replace("'", "''")
    foreach ($routine in @('CopiarBalances_BaseMes', 'CopiarBalances_BaseAnual', 'ActualizarSeriesSinPortapapeles')) {
        $excel.Run("'$escapedWorkbookName'!$routine")
    }

    try {
        $excel.CalculateUntilAsyncQueriesDone()
    } catch {
    }
    $excel.CalculateFullRebuild()

    $calculationDeadline = [DateTime]::UtcNow.AddMinutes(5)
    while ($excel.CalculationState -ne 0) {
        if ([DateTime]::UtcNow -ge $calculationDeadline) {
            throw 'Excel no terminó el recálculo de la plantilla dentro de cinco minutos.'
        }
        Start-Sleep -Milliseconds 500
    }

    $fechaGuardada = [DateTime]::FromOADate([double]$caratula.Range('B2').Value2)
    if ($fechaGuardada.Date -ne $fecha.Date) {
        throw "CARATULA!B2 no conserva la fecha solicitada. Esperado=$FechaCorte encontrado=$($fechaGuardada.ToString('yyyy-MM-dd'))"
    }

    $baseAnual = $workbook.Worksheets.Item('base anual')
    $ultimoRegistro = $baseAnual.Cells($baseAnual.Rows.Count, 2).End(-4162).Row
    $fechaBase = [DateTime]::new($fecha.Year, $fecha.Month, 1)
    $coincidenciasFecha = $excel.WorksheetFunction.CountIf($baseAnual.Range("B1:B$ultimoRegistro"), $fechaBase.ToOADate())
    if ([double]$coincidenciasFecha -le 0) {
        throw "La hoja 'base anual' no contiene información para $($fechaBase.ToString('yyyy-MM-dd')) después de ejecutar la macro."
    }

    $cuentas = $workbook.Worksheets.Item('cuentas')
    foreach ($cellRef in @('C4', 'C6', 'G15')) {
        $cell = $cuentas.Range($cellRef)
        if ([string]$cell.Text -match '^#') {
            throw "La celda cuentas!$cellRef contiene un error de Excel: $($cell.Text)"
        }
        try {
            [void][System.Convert]::ToDouble($cell.Value2, $culture)
        } catch {
            throw "La celda cuentas!$cellRef no contiene un valor numérico después de actualizar la plantilla."
        }
    }

    $workbook.Save()
    $saveChanges = $true
    Write-Output "PLANTILLA_ACTUALIZADA_OK ruta=$plantilla fecha=$FechaCorte rutinas=CopiarBalances_BaseMes,CopiarBalances_BaseAnual,ActualizarSeriesSinPortapapeles registrosBaseAnual=$([int]$coincidenciasFecha)"
} finally {
    if ($null -ne $workbook) {
        try { $workbook.Close($saveChanges) } catch {}
    }
    if ($null -ne $excel) {
        try { $excel.Quit() } catch {}
    }
    foreach ($comObject in @($cell, $cuentas, $baseAnual, $caratula, $workbook, $excel)) {
        if ($null -ne $comObject) {
            try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($comObject) } catch {}
        }
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
