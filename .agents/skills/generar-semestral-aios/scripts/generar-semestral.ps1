[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$FechaCorte,

    [string]$ProjectDir = 'D:\app\plantilla-aios-macro',
    [string]$PlantillaPath,
    [string]$OutputDir,
    [string]$MavenRepository,
    [int]$PreferredPort = 18085,
    [int]$MacroTimeoutMinutes = 30,
    [switch]$KeepApplicationRunning,
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

$lastDay = [DateTime]::DaysInMonth($fecha.Year, $fecha.Month)
if (@(6, 12) -notcontains $fecha.Month -or $fecha.Day -ne $lastDay) {
    throw "La fecha semestral debe ser 30 de junio o 31 de diciembre. Valor recibido: $FechaCorte"
}
if ($MacroTimeoutMinutes -lt 1) {
    throw 'MacroTimeoutMinutes debe ser al menos 1.'
}

$projectPath = [System.IO.Path]::GetFullPath($ProjectDir)
if (-not (Test-Path -LiteralPath (Join-Path $projectPath 'pom.xml') -PathType Leaf)) {
    throw "No se encontró pom.xml en $projectPath"
}

if ([string]::IsNullOrWhiteSpace($PlantillaPath)) {
    $PlantillaPath = Join-Path $projectPath 'plantillas\Plantilla AIOS-probable.xlsm'
}
$plantilla = [System.IO.Path]::GetFullPath($PlantillaPath)
if (-not (Test-Path -LiteralPath $plantilla -PathType Leaf)) {
    throw "No se encontró la plantilla AIOS: $plantilla"
}

$updateScript = Join-Path $PSScriptRoot 'actualizar-plantilla-aios.ps1'
if (-not (Test-Path -LiteralPath $updateScript -PathType Leaf)) {
    throw "No se encontró el script de actualización de la plantilla: $updateScript"
}

if ($SoloValidar) {
    & $updateScript -PlantillaPath $plantilla -FechaCorte $FechaCorte -SoloValidar
    $maven = Get-Command 'mvn.cmd' -ErrorAction SilentlyContinue
    if ($null -eq $maven) {
        $maven = Get-Command 'mvn' -ErrorAction Stop
    }
    Write-Output "VALIDACION_SEMESTRAL_OK proyecto=$projectPath excel=disponible maven=$($maven.Source)"
    return
}

if ([string]::IsNullOrWhiteSpace($OutputDir)) {
    $OutputDir = Join-Path $projectPath "target\aios-output\semestral-$($fecha.ToString('yyyy-MM'))"
}
$outputPath = [System.IO.Path]::GetFullPath($OutputDir)
New-Item -ItemType Directory -Path $outputPath -Force | Out-Null

if ([string]::IsNullOrWhiteSpace($MavenRepository)) {
    $MavenRepository = Join-Path $env:USERPROFILE '.m2\repository'
}
$mavenRepositoryPath = [System.IO.Path]::GetFullPath($MavenRepository)

function Quote-ProcessArgument {
    param([Parameter(Mandatory = $true)][string]$Value)
    return '"' + $Value.Replace('"', '\"') + '"'
}

function Get-MacroState {
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { return '' }
    return (Get-Content -LiteralPath $Path -Raw).Trim()
}

function Stop-MacroExcelProcess {
    param([string]$PidFile)
    if (-not (Test-Path -LiteralPath $PidFile -PathType Leaf)) { return }
    $excelPidText = (Get-Content -LiteralPath $PidFile -Raw).Trim()
    [int]$excelPid = 0
    if ([int]::TryParse($excelPidText, [ref]$excelPid)) {
        Stop-Process -Id $excelPid -Force -ErrorAction SilentlyContinue
    }
    Remove-Item -LiteralPath $PidFile -Force -ErrorAction SilentlyContinue
}

function Stop-TemporaryApplication {
    param(
        [System.Diagnostics.Process]$MavenProcess,
        [int]$Port
    )
    try {
        $listeners = @(Get-NetTCPConnection -LocalPort $Port -State Listen -ErrorAction SilentlyContinue)
        foreach ($listener in $listeners) {
            Stop-Process -Id $listener.OwningProcess -Force -ErrorAction SilentlyContinue
        }
    } catch {
    }
    if ($null -ne $MavenProcess -and -not $MavenProcess.HasExited) {
        Stop-Process -Id $MavenProcess.Id -Force -ErrorAction SilentlyContinue
    }
}

function Get-FreeLocalPort {
    param([int]$StartPort)
    foreach ($candidate in $StartPort..($StartPort + 100)) {
        $listener = $null
        try {
            $listener = [System.Net.Sockets.TcpListener]::new([System.Net.IPAddress]::Loopback, $candidate)
            $listener.Start()
            return $candidate
        } catch {
        } finally {
            if ($null -ne $listener) {
                try { $listener.Stop() } catch {}
            }
        }
    }
    throw "No se encontró un puerto libre entre $StartPort y $($StartPort + 100)."
}

function Test-AiosApplication {
    param([string]$BaseUrl)
    try {
        $response = Invoke-WebRequest -Uri "$BaseUrl/actuator/health" -Method Get -TimeoutSec 3 -UseBasicParsing
        return $response.StatusCode -eq 200
    } catch {
        return $false
    }
}

function Get-ZipEntryText {
    param(
        [System.IO.Compression.ZipArchive]$Archive,
        [string]$EntryName
    )
    $entry = $Archive.GetEntry($EntryName)
    if ($null -eq $entry) { return '' }
    $stream = $entry.Open()
    $reader = [System.IO.StreamReader]::new($stream)
    try { return $reader.ReadToEnd() } finally { $reader.Dispose(); $stream.Dispose() }
}

function Test-SemestralWorkbook {
    param(
        [string]$Path,
        [DateTime]$Cutoff
    )
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $archive = [System.IO.Compression.ZipFile]::OpenRead($Path)
    try {
        [xml]$workbookXml = Get-ZipEntryText -Archive $archive -EntryName 'xl/workbook.xml'
        $sheetNames = @($workbookXml.workbook.sheets.sheet | ForEach-Object { [string]$_.name })
        if ($sheetNames.Count -lt 1) {
            throw 'El archivo semestral no contiene hojas.'
        }

        $periodMonth = if ($Cutoff.Month -eq 6) { 'junio' } else { 'diciembre' }
        $periodYear = $Cutoff.ToString('yyyy')
        $xmlContent = [System.Text.StringBuilder]::new()
        foreach ($entry in $archive.Entries) {
            if ($entry.FullName -eq 'xl/sharedStrings.xml' -or $entry.FullName -like 'xl/worksheets/*.xml') {
                [void]$xmlContent.Append((Get-ZipEntryText -Archive $archive -EntryName $entry.FullName))
            }
        }
        $content = $xmlContent.ToString()
        if ($content.IndexOf($periodMonth, [StringComparison]::OrdinalIgnoreCase) -lt 0 -or $content.IndexOf($periodYear, [StringComparison]::OrdinalIgnoreCase) -lt 0) {
            throw "No se encontró el período $periodMonth $periodYear en el archivo semestral generado."
        }
        foreach ($formulaError in @('#REF!', '#DIV/0!', '#VALUE!', '#NAME?', '#N/A')) {
            if ($content.IndexOf($formulaError, [StringComparison]::OrdinalIgnoreCase) -ge 0) {
                throw "Se encontró el error de fórmula $formulaError en el archivo semestral generado."
            }
        }
        if ([regex]::Matches($content, '<v>-?\d+(?:\.\d+)?</v>').Count -lt 10) {
            throw 'El archivo semestral no contiene suficientes valores numéricos para considerarse generado correctamente.'
        }
        return "$periodMonth-$periodYear"
    } finally {
        $archive.Dispose()
    }
}

$backupDir = Join-Path $projectPath 'target\aios-backups'
New-Item -ItemType Directory -Path $backupDir -Force | Out-Null
$backupPath = Join-Path $backupDir ("Plantilla AIOS-probable.before-{0}-{1}.xlsm" -f $fecha.ToString('yyyy-MM-dd'), [DateTime]::Now.ToString('yyyyMMdd-HHmmssfff'))
Copy-Item -LiteralPath $plantilla -Destination $backupPath -ErrorAction Stop

$excelPidFile = Join-Path $outputPath '.excel-process.pid'
$macroStateFile = Join-Path $outputPath '.macro-state'
$macroStdout = Join-Path $outputPath 'actualizacion-plantilla.stdout.log'
$macroStderr = Join-Path $outputPath 'actualizacion-plantilla.stderr.log'
Remove-Item -LiteralPath $excelPidFile, $macroStateFile -Force -ErrorAction SilentlyContinue
$powershell = (Get-Command 'powershell.exe' -ErrorAction Stop).Source
$workerArguments = @(
    '-NoProfile',
    '-ExecutionPolicy', 'Bypass',
    '-File', (Quote-ProcessArgument -Value $updateScript),
    '-PlantillaPath', (Quote-ProcessArgument -Value $plantilla),
    '-FechaCorte', $FechaCorte,
    '-ExcelProcessIdFile', (Quote-ProcessArgument -Value $excelPidFile),
    '-MacroStateFile', (Quote-ProcessArgument -Value $macroStateFile)
)

$macroProcess = Start-Process -FilePath $powershell `
    -ArgumentList $workerArguments `
    -WindowStyle Hidden `
    -RedirectStandardOutput $macroStdout `
    -RedirectStandardError $macroStderr `
    -PassThru

$macroDeadline = [DateTime]::UtcNow.AddMinutes($MacroTimeoutMinutes)
$macroTimedOut = $false
while (-not $macroProcess.HasExited) {
    if ([DateTime]::UtcNow -ge $macroDeadline) {
        Stop-Process -Id $macroProcess.Id -Force -ErrorAction SilentlyContinue
        [void]$macroProcess.WaitForExit(5000)
        Stop-MacroExcelProcess -PidFile $excelPidFile
        $macroTimedOut = $true
        break
    }
    Start-Sleep -Seconds 2
}

$macroResult = if (Test-Path -LiteralPath $macroStdout) { Get-Content -LiteralPath $macroStdout -Raw } else { '' }
$macroError = if (Test-Path -LiteralPath $macroStderr) { Get-Content -LiteralPath $macroStderr -Raw } else { '' }
$macroState = Get-MacroState -Path $macroStateFile
$recoverableStates = @('PREPARACION_INICIADA', 'MACRO_INICIADA', 'OMITIDA')
$macroPreparationStatus = 'actualizada'
$macroWarning = ''

if ($macroTimedOut) {
    if ($recoverableStates -notcontains $macroState) {
        Remove-Item -LiteralPath $macroStateFile -Force -ErrorAction SilentlyContinue
        throw "La validación previa de la plantilla superó el límite de $MacroTimeoutMinutes minutos antes de iniciar la macro. Revise $macroStderr"
    }
    $macroPreparationStatus = 'omitida'
    $macroWarning = "ADVERTENCIA_MACRO_PLANTILLA_OMITIDA fecha=$FechaCorte motivo=tiempo_limite minutos=$MacroTimeoutMinutes accion=continuar_con_aplicacion stdout=$macroStdout stderr=$macroStderr"
} elseif ($macroProcess.ExitCode -ne 0) {
    if ($recoverableStates -notcontains $macroState) {
        Remove-Item -LiteralPath $excelPidFile, $macroStateFile -Force -ErrorAction SilentlyContinue
        throw "Falló la validación previa de Plantilla AIOS-probable.xlsm antes de iniciar la macro. No se generó el semestral. $macroError"
    }
    Stop-MacroExcelProcess -PidFile $excelPidFile
    $macroPreparationStatus = 'omitida'
    $macroWarning = "ADVERTENCIA_MACRO_PLANTILLA_OMITIDA fecha=$FechaCorte motivo=proceso_finalizado_con_error accion=continuar_con_aplicacion stdout=$macroStdout stderr=$macroStderr"
} elseif ($macroResult -match 'PLANTILLA_SEMESTRAL_PREPARADA_OK' -or $macroState -eq 'COMPLETADA') {
    $macroPreparationStatus = 'actualizada'
} elseif ($macroResult -match 'ADVERTENCIA_MACRO_PLANTILLA_OMITIDA' -or $macroState -eq 'OMITIDA' -or ($recoverableStates -contains $macroState)) {
    $macroPreparationStatus = 'omitida'
    $macroWarningReason = if ($macroResult -match 'ADVERTENCIA_MACRO_PLANTILLA_OMITIDA' -or $macroState -eq 'OMITIDA') { 'error_en_macro_o_validacion_posterior' } else { 'preparacion_no_confirmada' }
    $macroWarning = "ADVERTENCIA_MACRO_PLANTILLA_OMITIDA fecha=$FechaCorte motivo=$macroWarningReason accion=continuar_con_aplicacion stdout=$macroStdout stderr=$macroStderr"
} else {
    Remove-Item -LiteralPath $excelPidFile -Force -ErrorAction SilentlyContinue
    Remove-Item -LiteralPath $macroStateFile -Force -ErrorAction SilentlyContinue
    throw "La validación previa de la plantilla terminó sin confirmación verificable. Revise $macroStdout y $macroStderr"
}
Remove-Item -LiteralPath $excelPidFile, $macroStateFile -Force -ErrorAction SilentlyContinue

$mavenProcess = $null
$port = Get-FreeLocalPort -StartPort $PreferredPort
$baseUrl = "http://localhost:$port"
try {
    $maven = Get-Command 'mvn.cmd' -ErrorAction SilentlyContinue
    if ($null -eq $maven) {
        $maven = Get-Command 'mvn' -ErrorAction Stop
    }
    $appStdout = Join-Path $outputPath 'spring-boot.stdout.log'
    $appStderr = Join-Path $outputPath 'spring-boot.stderr.log'
    $mavenProcess = Start-Process -FilePath $maven.Source `
        -ArgumentList "-Dmaven.repo.local=$mavenRepositoryPath", '-DskipTests', 'spring-boot:run', "-Dspring-boot.run.arguments=--server.port=$port" `
        -WorkingDirectory $projectPath `
        -WindowStyle Hidden `
        -RedirectStandardOutput $appStdout `
        -RedirectStandardError $appStderr `
        -PassThru

    $applicationDeadline = [DateTime]::UtcNow.AddMinutes(3)
    while (-not (Test-AiosApplication -BaseUrl $baseUrl)) {
        if ($mavenProcess.HasExited) {
            throw "La aplicación terminó antes de estar disponible. Revise $appStderr"
        }
        if ([DateTime]::UtcNow -ge $applicationDeadline) {
            throw "La aplicación no respondió antes del tiempo límite. Revise $appStdout y $appStderr"
        }
        Start-Sleep -Seconds 2
    }

    $destination = Join-Path $outputPath 'semestral.xlsx'
    $uri = "$baseUrl/aios/generar?fechaCorte=$FechaCorte&modo=SEMESTRAL"
    Invoke-WebRequest -Uri $uri -Method Post -OutFile $destination -TimeoutSec 1800 -UseBasicParsing
    $file = Get-Item -LiteralPath $destination
    if ($file.Length -eq 0) {
        throw "La generación semestral de $FechaCorte produjo un archivo vacío."
    }
    $periodLabel = Test-SemestralWorkbook -Path $file.FullName -Cutoff $fecha
    if (-not [string]::IsNullOrWhiteSpace($macroResult)) {
        Write-Output $macroResult.Trim()
    }
    if (-not [string]::IsNullOrWhiteSpace($macroWarning)) {
        Write-Output $macroWarning
    }
    Write-Output "RESPALDO_PLANTILLA ruta=$backupPath"
    Write-Output "SEMESTRAL_GENERADO_OK fecha=$FechaCorte periodo=$periodLabel preparacionPlantilla=$macroPreparationStatus bytes=$($file.Length) ruta=$($file.FullName)"
    $file.FullName
} finally {
    if ($null -ne $mavenProcess -and -not $KeepApplicationRunning) {
        Stop-TemporaryApplication -MavenProcess $mavenProcess -Port $port
    }
    if (Test-Path -LiteralPath $excelPidFile) {
        Remove-Item -LiteralPath $excelPidFile -Force -ErrorAction SilentlyContinue
    }
    if (Test-Path -LiteralPath $macroStateFile) {
        Remove-Item -LiteralPath $macroStateFile -Force -ErrorAction SilentlyContinue
    }
}
