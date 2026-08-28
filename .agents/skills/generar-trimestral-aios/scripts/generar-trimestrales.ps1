[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$Desde,

    [string]$Hasta,
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

if ([string]::IsNullOrWhiteSpace($Hasta)) { $Hasta = $Desde }
$culture = [System.Globalization.CultureInfo]::InvariantCulture
$styles = [System.Globalization.DateTimeStyles]::None
$inicio = [DateTime]::MinValue
$fin = [DateTime]::MinValue
if (-not [DateTime]::TryParseExact($Desde, 'yyyy-MM', $culture, $styles, [ref]$inicio)) {
    throw "Desde debe tener formato AAAA-MM. Valor recibido: $Desde"
}
if (-not [DateTime]::TryParseExact($Hasta, 'yyyy-MM', $culture, $styles, [ref]$fin)) {
    throw "Hasta debe tener formato AAAA-MM. Valor recibido: $Hasta"
}
if ($inicio -gt $fin) { throw "Desde ($Desde) no puede ser posterior a Hasta ($Hasta)." }

$fechasCorte = [System.Collections.Generic.List[DateTime]]::new()
$actual = [DateTime]::new($inicio.Year, $inicio.Month, 1)
$ultimo = [DateTime]::new($fin.Year, $fin.Month, 1)
while ($actual -le $ultimo) {
    if (@(3, 6, 9, 12) -contains $actual.Month) {
        $fechasCorte.Add([DateTime]::new($actual.Year, $actual.Month, [DateTime]::DaysInMonth($actual.Year, $actual.Month)))
    }
    $actual = $actual.AddMonths(1)
}
if ($fechasCorte.Count -eq 0) {
    throw "El rango $Desde a $Hasta no contiene cortes trimestrales."
}

$projectPath = [System.IO.Path]::GetFullPath($ProjectDir)
if (-not (Test-Path -LiteralPath (Join-Path $projectPath 'pom.xml') -PathType Leaf)) {
    throw "No se encontró pom.xml en $projectPath"
}
$singleScript = Join-Path $PSScriptRoot 'generar-trimestral.ps1'
if (-not (Test-Path -LiteralPath $singleScript -PathType Leaf)) {
    throw "No se encontró el generador trimestral individual: $singleScript"
}
if ([string]::IsNullOrWhiteSpace($OutputDir)) {
    $OutputDir = Join-Path $projectPath "target\aios-output\trimestrales-$Desde-a-$Hasta"
}
$outputPath = [System.IO.Path]::GetFullPath($OutputDir)
New-Item -ItemType Directory -Path $outputPath -Force | Out-Null

foreach ($fecha in $fechasCorte) {
    $fechaTexto = $fecha.ToString('yyyy-MM-dd')
    $params = @{
        FechaCorte = $fechaTexto
        ProjectDir = $projectPath
        OutputDir = (Join-Path $outputPath "preparacion-$fechaTexto")
        PreferredPort = $PreferredPort
        MacroTimeoutMinutes = $MacroTimeoutMinutes
    }
    if (-not [string]::IsNullOrWhiteSpace($PlantillaPath)) { $params.PlantillaPath = $PlantillaPath }
    if (-not [string]::IsNullOrWhiteSpace($MavenRepository)) { $params.MavenRepository = $MavenRepository }
    if ($SoloValidar) { $params.SoloValidar = $true } else { $params.SoloPreparar = $true }
    & $singleScript @params
}
if ($SoloValidar) {
    Write-Output "VALIDACION_TRIMESTRALES_OK desde=$Desde hasta=$Hasta periodos=$($fechasCorte.Count)"
    return
}

if ([string]::IsNullOrWhiteSpace($MavenRepository)) {
    $MavenRepository = Join-Path $env:USERPROFILE '.m2\repository'
}
$mavenRepositoryPath = [System.IO.Path]::GetFullPath($MavenRepository)

function Stop-TemporaryApplication {
    param([System.Diagnostics.Process]$MavenProcess, [int]$Port)
    try {
        foreach ($listener in @(Get-NetTCPConnection -LocalPort $Port -State Listen -ErrorAction SilentlyContinue)) {
            Stop-Process -Id $listener.OwningProcess -Force -ErrorAction SilentlyContinue
        }
    } catch {}
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
        } catch {} finally { if ($null -ne $listener) { try { $listener.Stop() } catch {} } }
    }
    throw "No se encontró un puerto libre entre $StartPort y $($StartPort + 100)."
}

function Test-AiosApplication {
    param([string]$BaseUrl)
    try {
        $response = Invoke-WebRequest -Uri "$BaseUrl/actuator/health" -Method Get -TimeoutSec 3 -UseBasicParsing
        return $response.StatusCode -eq 200
    } catch { return $false }
}

function Get-ZipEntryText {
    param([System.IO.Compression.ZipArchive]$Archive, [string]$EntryName)
    $entry = $Archive.GetEntry($EntryName)
    if ($null -eq $entry) { return '' }
    $stream = $entry.Open()
    $reader = [System.IO.StreamReader]::new($stream)
    try { return $reader.ReadToEnd() } finally { $reader.Dispose(); $stream.Dispose() }
}

function Test-TrimestralesWorkbook {
    param([string]$Path, [System.Collections.Generic.List[DateTime]]$Cutoffs)
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $archive = [System.IO.Compression.ZipFile]::OpenRead($Path)
    try {
        [xml]$workbookXml = Get-ZipEntryText -Archive $archive -EntryName 'xl/workbook.xml'
        $sheetNames = @($workbookXml.workbook.sheets.sheet | ForEach-Object { [string]$_.name })
        $expectedSheets = @('afiliados', 'aportantes', 'colombia', 'gastos', 'comisiones', 'rentabilidad', 'promotores', 'traspasos')
        $missing = @($expectedSheets | Where-Object { $sheetNames -notcontains $_ })
        if ($missing.Count -gt 0) { throw "Faltan hojas trimestrales: $($missing -join ', ')" }

        $xmlContent = [System.Text.StringBuilder]::new()
        foreach ($entry in $archive.Entries) {
            if ($entry.FullName -eq 'xl/sharedStrings.xml' -or $entry.FullName -like 'xl/worksheets/*.xml') {
                [void]$xmlContent.Append((Get-ZipEntryText -Archive $archive -EntryName $entry.FullName))
            }
        }
        $content = $xmlContent.ToString()
        $months = @('', 'ene', 'feb', 'mar', 'abr', 'may', 'jun', 'jul', 'ago', 'sep', 'oct', 'nov', 'dic')
        $labels = foreach ($cutoff in $Cutoffs) { "$($months[$cutoff.Month])-$($cutoff.ToString('yy'))" }
        foreach ($label in $labels) {
            if ($content.IndexOf($label, [StringComparison]::OrdinalIgnoreCase) -lt 0) {
                throw "No se encontró la etiqueta $label en el trimestral consolidado."
            }
        }
        foreach ($formulaError in @('#REF!', '#DIV/0!', '#VALUE!', '#NAME?', '#N/A')) {
            if ($content.IndexOf($formulaError, [StringComparison]::OrdinalIgnoreCase) -ge 0) {
                throw "Se encontró el error de fórmula $formulaError en el trimestral consolidado."
            }
        }
        return ($labels -join ',')
    } finally { $archive.Dispose() }
}

$mavenProcess = $null
$port = Get-FreeLocalPort -StartPort $PreferredPort
$baseUrl = "http://localhost:$port"
try {
    $maven = Get-Command 'mvn.cmd' -ErrorAction SilentlyContinue
    if ($null -eq $maven) { $maven = Get-Command 'mvn' -ErrorAction Stop }
    $appStdout = Join-Path $outputPath 'spring-boot.stdout.log'
    $appStderr = Join-Path $outputPath 'spring-boot.stderr.log'
    $mavenProcess = Start-Process -FilePath $maven.Source `
        -ArgumentList "-Dmaven.repo.local=$mavenRepositoryPath", '-DskipTests', 'spring-boot:run', "-Dspring-boot.run.arguments=--server.port=$port" `
        -WorkingDirectory $projectPath -WindowStyle Hidden `
        -RedirectStandardOutput $appStdout -RedirectStandardError $appStderr -PassThru

    $deadline = [DateTime]::UtcNow.AddMinutes(3)
    while (-not (Test-AiosApplication -BaseUrl $baseUrl)) {
        if ($mavenProcess.HasExited) { throw "La aplicación terminó antes de estar disponible. Revise $appStderr" }
        if ([DateTime]::UtcNow -ge $deadline) { throw "La aplicación no respondió antes del tiempo límite. Revise $appStdout y $appStderr" }
        Start-Sleep -Seconds 2
    }

    $desdeCorte = [DateTime]::new($inicio.Year, $inicio.Month, 1).ToString('yyyy-MM-dd')
    $hastaCorte = [DateTime]::new($fin.Year, $fin.Month, [DateTime]::DaysInMonth($fin.Year, $fin.Month)).ToString('yyyy-MM-dd')
    $destination = Join-Path $outputPath 'Boletin_AIOS TRIMESTRAL.xlsx'
    $uri = "$baseUrl/aios/generar-rango?desde=$desdeCorte&hasta=$hastaCorte&modo=TRIMESTRAL"
    Invoke-WebRequest -Uri $uri -Method Post -OutFile $destination -TimeoutSec 3600 -UseBasicParsing
    $file = Get-Item -LiteralPath $destination
    if ($file.Length -eq 0) { throw "La generación trimestral de $Desde a $Hasta produjo un archivo vacío." }
    $periodLabels = Test-TrimestralesWorkbook -Path $file.FullName -Cutoffs $fechasCorte
    Write-Output "TRIMESTRALES_GENERADOS_OK desde=$Desde hasta=$Hasta periodos=$periodLabels bytes=$($file.Length) ruta=$($file.FullName)"
    $file.FullName
} finally {
    if ($null -ne $mavenProcess -and -not $KeepApplicationRunning) {
        Stop-TemporaryApplication -MavenProcess $mavenProcess -Port $port
    }
}
