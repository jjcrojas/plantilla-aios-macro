[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$FechaCorte,

    [string]$ProjectDir = 'D:\app\plantilla-aios-macro',
    [string]$PlantillaPath,
    [string]$OutputDir,
    [string]$MavenRepository,
    [int]$PreferredPort = 18084,
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

$quarterMonths = @(3, 6, 9, 12)
$lastDay = [DateTime]::DaysInMonth($fecha.Year, $fecha.Month)
if ($quarterMonths -notcontains $fecha.Month -or $fecha.Day -ne $lastDay) {
    throw "La fecha trimestral debe ser 31 de marzo, 30 de junio, 30 de septiembre o 31 de diciembre. Valor recibido: $FechaCorte"
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
    Write-Output "VALIDACION_TRIMESTRAL_OK proyecto=$projectPath excel=disponible maven=$($maven.Source)"
    return
}

if ([string]::IsNullOrWhiteSpace($OutputDir)) {
    $OutputDir = Join-Path $projectPath "target\aios-output\trimestral-$($fecha.ToString('yyyy-MM'))"
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

function Test-TrimestralWorkbook {
    param(
        [string]$Path,
        [DateTime]$Cutoff
    )
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $archive = [System.IO.Compression.ZipFile]::OpenRead($Path)
    try {
        [xml]$workbookXml = Get-ZipEntryText -Archive $archive -EntryName 'xl/workbook.xml'
        $sheetNames = @($workbookXml.workbook.sheets.sheet | ForEach-Object { [string]$_.name })
        $expected = @('afiliados', 'aportantes', 'colombia', 'gastos', 'comisiones', 'rentabilidad', 'promotores', 'traspasos')
        $missing = @($expected | Where-Object { $sheetNames -notcontains $_ })
        if ($missing.Count -gt 0) {
            throw "El archivo trimestral no contiene las hojas requeridas: $($missing -join ', ')"
        }

        $monthLabels = @('', 'ene', 'feb', 'mar', 'abr', 'may', 'jun', 'jul', 'ago', 'sep', 'oct', 'nov', 'dic')
        $periodLabel = "$($monthLabels[$Cutoff.Month])-$($Cutoff.ToString('yy'))"
        $xmlContent = [System.Text.StringBuilder]::new()
        foreach ($entry in $archive.Entries) {
            if ($entry.FullName -eq 'xl/sharedStrings.xml' -or $entry.FullName -like 'xl/worksheets/*.xml') {
                [void]$xmlContent.Append((Get-ZipEntryText -Archive $archive -EntryName $entry.FullName))
            }
        }
        if ($xmlContent.ToString().IndexOf($periodLabel, [StringComparison]::OrdinalIgnoreCase) -lt 0) {
            throw "No se encontró la etiqueta $periodLabel en el archivo trimestral generado."
        }
        foreach ($formulaError in @('#REF!', '#DIV/0!', '#VALUE!', '#NAME?', '#N/A')) {
            if ($xmlContent.ToString().IndexOf($formulaError, [StringComparison]::OrdinalIgnoreCase) -ge 0) {
                throw "Se encontró el error de fórmula $formulaError en el archivo trimestral generado."
            }
        }
        return $periodLabel
    } finally {
        $archive.Dispose()
    }
}

$backupDir = Join-Path $projectPath 'target\aios-backups'
New-Item -ItemType Directory -Path $backupDir -Force | Out-Null
$backupPath = Join-Path $backupDir ("Plantilla AIOS-probable.before-{0}-{1}.xlsm" -f $fecha.ToString('yyyy-MM-dd'), [DateTime]::Now.ToString('yyyyMMdd-HHmmssfff'))
Copy-Item -LiteralPath $plantilla -Destination $backupPath -ErrorAction Stop

$excelPidFile = Join-Path $outputPath '.excel-process.pid'
$macroStdout = Join-Path $outputPath 'actualizacion-plantilla.stdout.log'
$macroStderr = Join-Path $outputPath 'actualizacion-plantilla.stderr.log'
$powershell = (Get-Command 'powershell.exe' -ErrorAction Stop).Source
$workerArguments = @(
    '-NoProfile',
    '-ExecutionPolicy', 'Bypass',
    '-File', (Quote-ProcessArgument -Value $updateScript),
    '-PlantillaPath', (Quote-ProcessArgument -Value $plantilla),
    '-FechaCorte', $FechaCorte,
    '-ExcelProcessIdFile', (Quote-ProcessArgument -Value $excelPidFile)
)

$macroProcess = Start-Process -FilePath $powershell `
    -ArgumentList $workerArguments `
    -WindowStyle Hidden `
    -RedirectStandardOutput $macroStdout `
    -RedirectStandardError $macroStderr `
    -PassThru

$macroDeadline = [DateTime]::UtcNow.AddMinutes($MacroTimeoutMinutes)
while (-not $macroProcess.HasExited) {
    if ([DateTime]::UtcNow -ge $macroDeadline) {
        Stop-Process -Id $macroProcess.Id -Force -ErrorAction SilentlyContinue
        if (Test-Path -LiteralPath $excelPidFile) {
            $excelPidText = (Get-Content -LiteralPath $excelPidFile -Raw).Trim()
            [int]$excelPid = 0
            if ([int]::TryParse($excelPidText, [ref]$excelPid)) {
                Stop-Process -Id $excelPid -Force -ErrorAction SilentlyContinue
            }
            Remove-Item -LiteralPath $excelPidFile -Force -ErrorAction SilentlyContinue
        }
        throw "La actualización contable de la plantilla superó el límite de $MacroTimeoutMinutes minutos. Revise $macroStderr"
    }
    Start-Sleep -Seconds 2
}

if ($macroProcess.ExitCode -ne 0) {
    $macroError = if (Test-Path -LiteralPath $macroStderr) { Get-Content -LiteralPath $macroStderr -Raw } else { '' }
    Remove-Item -LiteralPath $excelPidFile -Force -ErrorAction SilentlyContinue
    throw "Falló la actualización de Plantilla AIOS-probable.xlsm. No se generó el trimestral. $macroError"
}
$macroResult = if (Test-Path -LiteralPath $macroStdout) { Get-Content -LiteralPath $macroStdout -Raw } else { '' }
if ($macroResult -notmatch 'PLANTILLA_ACTUALIZADA_OK') {
    Remove-Item -LiteralPath $excelPidFile -Force -ErrorAction SilentlyContinue
    throw "La actualización de la plantilla terminó sin confirmación verificable. Revise $macroStdout"
}

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

    $destination = Join-Path $outputPath 'Boletin_AIOS TRIMESTRAL.xlsx'
    $uri = "$baseUrl/aios/generar?fechaCorte=$FechaCorte&modo=TRIMESTRAL"
    Invoke-WebRequest -Uri $uri -Method Post -OutFile $destination -TimeoutSec 1800 -UseBasicParsing
    $file = Get-Item -LiteralPath $destination
    if ($file.Length -eq 0) {
        throw "La generación trimestral de $FechaCorte produjo un archivo vacío."
    }
    $periodLabel = Test-TrimestralWorkbook -Path $file.FullName -Cutoff $fecha
    Write-Output $macroResult.Trim()
    Write-Output "RESPALDO_PLANTILLA ruta=$backupPath"
    Write-Output "TRIMESTRAL_GENERADO_OK fecha=$FechaCorte periodo=$periodLabel bytes=$($file.Length) ruta=$($file.FullName)"
    $file.FullName
} finally {
    if ($null -ne $mavenProcess -and -not $KeepApplicationRunning) {
        Stop-TemporaryApplication -MavenProcess $mavenProcess -Port $port
    }
    if (Test-Path -LiteralPath $excelPidFile) {
        Remove-Item -LiteralPath $excelPidFile -Force -ErrorAction SilentlyContinue
    }
}
