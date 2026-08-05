[CmdletBinding()]
param(
    [string]$ProjectDir = 'D:\app\plantilla-aios-macro',
    [string]$Desde,
    [string]$Hasta,
    [string]$OutputDir,
    [string]$BaseUrl = 'http://localhost:8084',
    [string]$MavenRepository,
    [switch]$KeepApplicationRunning
)

$ErrorActionPreference = 'Stop'

if ([string]::IsNullOrWhiteSpace($Desde) -and [string]::IsNullOrWhiteSpace($Hasta)) {
    $Desde = '2025-06'
    $Hasta = '2025-12'
} elseif ([string]::IsNullOrWhiteSpace($Desde)) {
    $Desde = $Hasta
} elseif ([string]::IsNullOrWhiteSpace($Hasta)) {
    $Hasta = $Desde
}

$periodFormat = 'yyyy-MM'
$culture = [System.Globalization.CultureInfo]::InvariantCulture
$styles = [System.Globalization.DateTimeStyles]::None
$startPeriod = [DateTime]::MinValue
$endPeriod = [DateTime]::MinValue
if (-not [DateTime]::TryParseExact($Desde, $periodFormat, $culture, $styles, [ref]$startPeriod)) {
    throw "Desde debe tener formato AAAA-MM. Valor recibido: $Desde"
}
if (-not [DateTime]::TryParseExact($Hasta, $periodFormat, $culture, $styles, [ref]$endPeriod)) {
    throw "Hasta debe tener formato AAAA-MM. Valor recibido: $Hasta"
}
if ($startPeriod -gt $endPeriod) {
    throw "Desde ($Desde) no puede ser posterior a Hasta ($Hasta)."
}

$projectPath = [System.IO.Path]::GetFullPath($ProjectDir)
if (-not (Test-Path -LiteralPath (Join-Path $projectPath 'pom.xml') -PathType Leaf)) {
    throw "No se encontró pom.xml en $projectPath"
}

if ([string]::IsNullOrWhiteSpace($OutputDir)) {
    $rangeLabel = if ($Desde -eq $Hasta) { "mensual-$Desde" } else { "mensuales-$Desde-a-$Hasta" }
    $OutputDir = Join-Path $projectPath "target\aios-output\$rangeLabel"
}
$outputPath = [System.IO.Path]::GetFullPath($OutputDir)
New-Item -ItemType Directory -Path $outputPath -Force | Out-Null

if ([string]::IsNullOrWhiteSpace($MavenRepository)) {
    $MavenRepository = Join-Path $env:USERPROFILE '.m2\repository'
}
$mavenRepositoryPath = [System.IO.Path]::GetFullPath($MavenRepository)

function Test-AiosApplication {
    param([string]$Url)
    try {
        $response = Invoke-WebRequest -Uri "$Url/actuator/health" -Method Get -TimeoutSec 3 -UseBasicParsing
        return $response.StatusCode -eq 200
    } catch {
        return $false
    }
}

$startedProcess = $null
try {
    if (-not (Test-AiosApplication -Url $BaseUrl)) {
        $maven = Get-Command 'mvn.cmd' -ErrorAction SilentlyContinue
        if ($null -eq $maven) {
            $maven = Get-Command 'mvn' -ErrorAction Stop
        }
        $stdoutLog = Join-Path $outputPath 'spring-boot.stdout.log'
        $stderrLog = Join-Path $outputPath 'spring-boot.stderr.log'
        $startedProcess = Start-Process -FilePath $maven.Source `
            -ArgumentList "-Dmaven.repo.local=$mavenRepositoryPath", '-DskipTests', 'spring-boot:run' `
            -WorkingDirectory $projectPath `
            -WindowStyle Hidden `
            -RedirectStandardOutput $stdoutLog `
            -RedirectStandardError $stderrLog `
            -PassThru

        $deadline = [DateTime]::UtcNow.AddMinutes(3)
        while (-not (Test-AiosApplication -Url $BaseUrl)) {
            if ($startedProcess.HasExited) {
                throw "La aplicación terminó antes de estar disponible. Revise $stderrLog"
            }
            if ([DateTime]::UtcNow -ge $deadline) {
                throw "La aplicación no respondió antes del tiempo límite. Revise $stdoutLog y $stderrLog"
            }
            Start-Sleep -Seconds 2
        }
    }

    $firstDay = [DateTime]::new($startPeriod.Year, $startPeriod.Month, 1)
    $lastDayNumber = [DateTime]::DaysInMonth($endPeriod.Year, $endPeriod.Month)
    $lastDay = [DateTime]::new($endPeriod.Year, $endPeriod.Month, $lastDayNumber)
    $desdeCorte = $firstDay.ToString('yyyy-MM-dd')
    $hastaCorte = $lastDay.ToString('yyyy-MM-dd')
    $destination = Join-Path $outputPath 'Boletin_AIOS MENSUAL.xlsx'
    $uri = "$BaseUrl/aios/generar-mensuales?desde=$desdeCorte&hasta=$hastaCorte"

    Invoke-WebRequest -Uri $uri -Method Post -OutFile $destination -TimeoutSec 1800 -UseBasicParsing
    $file = Get-Item -LiteralPath $destination
    if ($file.Length -eq 0) {
        throw "La generación de $Desde a $Hasta produjo un archivo vacío."
    }
    Write-Output "Generado $Desde a $Hasta -> $($file.FullName)"
} finally {
    if ($null -ne $startedProcess -and -not $KeepApplicationRunning -and -not $startedProcess.HasExited) {
        Stop-Process -Id $startedProcess.Id -Force
    }
}

Write-Output 'Total generado: 1 archivo consolidado.'
$file.FullName
