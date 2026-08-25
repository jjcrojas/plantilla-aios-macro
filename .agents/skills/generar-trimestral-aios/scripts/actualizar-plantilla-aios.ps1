[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$PlantillaPath,

    [Parameter(Mandatory = $true)]
    [string]$FechaCorte,

    [string]$ExcelProcessIdFile,

    [string]$MacroStateFile,

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
$baseMes = $null
$cuentas = $null
$cell = $null
$saveChanges = $false

function Set-MacroState {
    param([Parameter(Mandatory = $true)][string]$State)
    if (-not [string]::IsNullOrWhiteSpace($MacroStateFile)) {
        [System.IO.File]::WriteAllText([System.IO.Path]::GetFullPath($MacroStateFile), $State)
    }
}

function ConvertTo-SingleLine {
    param([string]$Value)
    if ([string]::IsNullOrWhiteSpace($Value)) { return 'sin detalle' }
    return ($Value -replace '[\r\n]+', ' ').Trim()
}

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

    $fechaBase = [DateTime]::new($fecha.Year, $fecha.Month, 1)
    $baseAnual = $workbook.Worksheets.Item('base anual')
    $ultimoRegistroAnual = $baseAnual.Cells($baseAnual.Rows.Count, 2).End(-4162).Row
    $coincidenciasAnual = $excel.WorksheetFunction.CountIf($baseAnual.Range("B1:B$ultimoRegistroAnual"), $fechaBase.ToOADate())
    if ([double]$coincidenciasAnual -le 0) {
        throw "La hoja 'base anual' no contiene información previamente actualizada para $($fechaBase.ToString('yyyy-MM-dd')). La skill no modifica esta hoja."
    }

    $baseMes = $workbook.Worksheets.Item('base mes')
    $ultimoRegistroMes = $baseMes.Cells($baseMes.Rows.Count, 4).End(-4162).Row
    $coincidenciasMesInicio = $excel.WorksheetFunction.CountIf($baseMes.Range("D1:D$ultimoRegistroMes"), $fechaBase.ToOADate())
    $coincidenciasMesCorte = $excel.WorksheetFunction.CountIf($baseMes.Range("D1:D$ultimoRegistroMes"), $fecha.ToOADate())
    $coincidenciasMes = [double]$coincidenciasMesInicio + [double]$coincidenciasMesCorte
    if ([double]$coincidenciasMes -le 0) {
        throw "La hoja 'base mes' no contiene información previamente actualizada para $($fechaBase.ToString('yyyy-MM-dd')) ni $FechaCorte. La skill no modifica esta hoja."
    }

    if ($SoloValidar) {
        Write-Output "VALIDACION_PLANTILLA_OK ruta=$plantilla macro=$macroAction fechaSolicitada=$FechaCorte registrosBaseAnual=$([int]$coincidenciasAnual) registrosBaseMes=$([int]$coincidenciasMes) basesModificadas=no"
        return
    }

    Set-MacroState -State 'PREPARACION_INICIADA'
    try {
        $caratula.Range('B2').Value2 = $fecha.ToOADate()
        $caratula.Range('B2').NumberFormat = 'dd/mm/yyyy'
        $excel.CalculateFull()

        $escapedWorkbookName = $workbook.Name.Replace("'", "''")
        Set-MacroState -State 'MACRO_INICIADA'
        Write-Output "MACRO_PLANTILLA_INICIADA ruta=$plantilla fecha=$FechaCorte rutina=ActualizarSeriesSinPortapapeles"
        $excel.Run("'$escapedWorkbookName'!ActualizarSeriesSinPortapapeles")

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
        Set-MacroState -State 'COMPLETADA'
        Write-Output "PLANTILLA_PREPARADA_OK ruta=$plantilla fecha=$FechaCorte rutina=ActualizarSeriesSinPortapapeles registrosBaseAnual=$([int]$coincidenciasAnual) registrosBaseMes=$([int]$coincidenciasMes) basesModificadas=no"
    } catch {
        $macroError = ConvertTo-SingleLine -Value $_.Exception.Message
        Set-MacroState -State 'OMITIDA'
        Write-Output "ADVERTENCIA_MACRO_PLANTILLA_OMITIDA fecha=$FechaCorte error=$macroError accion=continuar_con_aplicacion plantillaGuardada=no"
    }
} finally {
    if ($null -ne $workbook) {
        try { $workbook.Close($saveChanges) } catch {}
    }
    if ($null -ne $excel) {
        try { $excel.Quit() } catch {}
    }
    foreach ($comObject in @($cell, $cuentas, $baseMes, $baseAnual, $caratula, $workbook, $excel)) {
        if ($null -ne $comObject) {
            try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($comObject) } catch {}
        }
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
