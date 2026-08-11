[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$Desde,
    [string]$Hasta,
    [string]$ProjectDir = 'D:\app\plantilla-aios-macro',
    [string]$OutputDir,
    [string]$MavenRepository,
    [int]$PreferredPort = 18084,
    [int]$MacroTimeoutMinutes = 30,
    [switch]$SoloPlan
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

$meses = [System.Collections.Generic.List[DateTime]]::new()
$actual = [DateTime]::new($inicio.Year, $inicio.Month, 1)
$ultimo = [DateTime]::new($fin.Year, $fin.Month, 1)
while ($actual -le $ultimo) {
    $meses.Add($actual)
    $actual = $actual.AddMonths(1)
}
$trimestres = @($meses | Where-Object { $_.Month -in @(3, 6, 9, 12) } | ForEach-Object {
    [DateTime]::new($_.Year, $_.Month, [DateTime]::DaysInMonth($_.Year, $_.Month))
})
$semestres = @($meses | Where-Object { $_.Month -in @(6, 12) } | ForEach-Object {
    [DateTime]::new($_.Year, $_.Month, [DateTime]::DaysInMonth($_.Year, $_.Month))
})

$quarterLabels = @($trimestres | ForEach-Object { $_.ToString('yyyy-MM-dd') })
$semesterLabels = @($semestres | ForEach-Object { $_.ToString('yyyy-MM-dd') })
Write-Output "PLAN_MENSUAL desde=$Desde hasta=$Hasta meses=$($meses.Count)"
Write-Output "PLAN_TRIMESTRAL cortes=$($quarterLabels -join ',')"
Write-Output "PLAN_SEMESTRAL cortes=$($semesterLabels -join ',')"
if ($SoloPlan) { return }

$projectPath = [System.IO.Path]::GetFullPath($ProjectDir)
if (-not (Test-Path -LiteralPath (Join-Path $projectPath 'pom.xml') -PathType Leaf)) {
    throw "No se encontró pom.xml en $projectPath"
}
if ([string]::IsNullOrWhiteSpace($OutputDir)) {
    $label = if ($Desde -eq $Hasta) { "todos-$Desde" } else { "todos-$Desde-a-$Hasta" }
    $OutputDir = Join-Path $projectPath "target\aios-output\$label"
}
$outputPath = [System.IO.Path]::GetFullPath($OutputDir)
New-Item -ItemType Directory -Path $outputPath -Force | Out-Null

$monthlyScript = Join-Path $projectPath '.agents\skills\generar-mensuales-aios\scripts\generar-mensuales.ps1'
$quarterlyScript = Join-Path $projectPath '.agents\skills\generar-trimestral-aios\scripts\generar-trimestral.ps1'
$semesterScript = Join-Path $projectPath '.agents\skills\generar-semestral-aios\scripts\generar-semestral.ps1'
foreach ($script in @($monthlyScript, $quarterlyScript, $semesterScript)) {
    if (-not (Test-Path -LiteralPath $script -PathType Leaf)) { throw "No se encontró el script requerido: $script" }
}

$monthlyDir = Join-Path $outputPath 'mensual'
$monthlyArgs = @{ Desde = $Desde; Hasta = $Hasta; ProjectDir = $projectPath; OutputDir = $monthlyDir }
if (-not [string]::IsNullOrWhiteSpace($MavenRepository)) { $monthlyArgs.MavenRepository = $MavenRepository }
$port8084WasActive = $null -ne (Get-NetTCPConnection -LocalPort 8084 -State Listen -ErrorAction SilentlyContinue | Select-Object -First 1)
try {
    & $monthlyScript @monthlyArgs
} finally {
    if (-not $port8084WasActive) {
        $listeners = @(Get-NetTCPConnection -LocalPort 8084 -State Listen -ErrorAction SilentlyContinue)
        foreach ($listener in $listeners) { Stop-Process -Id $listener.OwningProcess -Force -ErrorAction SilentlyContinue }
    }
}
$monthlyFile = Get-Item -LiteralPath (Join-Path $monthlyDir 'Boletin_AIOS MENSUAL.xlsx')
Copy-Item -LiteralPath $monthlyFile.FullName -Destination (Join-Path $outputPath 'Boletin_AIOS MENSUAL.xlsx') -Force

$sourceReferences = Join-Path $projectPath 'salidas_referencia'
$previousReferenceEnv = $env:AIOS_SALIDAS_REFERENCIA_DIR
try {
    if ($trimestres.Count -gt 0) {
        $quarterReferenceDir = Join-Path $outputPath '.referencias-trimestral'
        New-Item -ItemType Directory -Path $quarterReferenceDir -Force | Out-Null
        Copy-Item -LiteralPath (Join-Path $sourceReferences 'Boletin_AIOS TRIMESTRAL.xlsx') -Destination $quarterReferenceDir -Force
        $env:AIOS_SALIDAS_REFERENCIA_DIR = $quarterReferenceDir
        for ($i = 0; $i -lt $trimestres.Count; $i++) {
            $cutoff = $trimestres[$i]
            $stepDir = Join-Path $outputPath "trimestral-$($cutoff.ToString('yyyy-MM'))"
            $args = @{ FechaCorte = $cutoff.ToString('yyyy-MM-dd'); ProjectDir = $projectPath; OutputDir = $stepDir; PreferredPort = ($PreferredPort + $i); MacroTimeoutMinutes = $MacroTimeoutMinutes }
            if (-not [string]::IsNullOrWhiteSpace($MavenRepository)) { $args.MavenRepository = $MavenRepository }
            & $quarterlyScript @args
            $stepFile = Get-Item -LiteralPath (Join-Path $stepDir 'Boletin_AIOS TRIMESTRAL.xlsx')
            Copy-Item -LiteralPath $stepFile.FullName -Destination (Join-Path $quarterReferenceDir 'Boletin_AIOS TRIMESTRAL.xlsx') -Force
        }
        Copy-Item -LiteralPath (Join-Path $quarterReferenceDir 'Boletin_AIOS TRIMESTRAL.xlsx') -Destination (Join-Path $outputPath 'Boletin_AIOS TRIMESTRAL.xlsx') -Force
    }

    if ($semestres.Count -gt 0) {
        $semesterReferenceDir = Join-Path $outputPath '.referencias-semestral'
        New-Item -ItemType Directory -Path $semesterReferenceDir -Force | Out-Null
        Copy-Item -LiteralPath (Join-Path $sourceReferences 'semestral.xlsx') -Destination $semesterReferenceDir -Force
        $env:AIOS_SALIDAS_REFERENCIA_DIR = $semesterReferenceDir
        for ($i = 0; $i -lt $semestres.Count; $i++) {
            $cutoff = $semestres[$i]
            $stepDir = Join-Path $outputPath "semestral-$($cutoff.ToString('yyyy-MM'))"
            $args = @{ FechaCorte = $cutoff.ToString('yyyy-MM-dd'); ProjectDir = $projectPath; OutputDir = $stepDir; PreferredPort = ($PreferredPort + 50 + $i); MacroTimeoutMinutes = $MacroTimeoutMinutes }
            if (-not [string]::IsNullOrWhiteSpace($MavenRepository)) { $args.MavenRepository = $MavenRepository }
            & $semesterScript @args
            $stepFile = Get-Item -LiteralPath (Join-Path $stepDir 'semestral.xlsx')
            Copy-Item -LiteralPath $stepFile.FullName -Destination (Join-Path $semesterReferenceDir 'semestral.xlsx') -Force
        }
        Copy-Item -LiteralPath (Join-Path $semesterReferenceDir 'semestral.xlsx') -Destination (Join-Path $outputPath 'semestral.xlsx') -Force
    }
} finally {
    $env:AIOS_SALIDAS_REFERENCIA_DIR = $previousReferenceEnv
}

$finalFiles = @('Boletin_AIOS MENSUAL.xlsx', 'Boletin_AIOS TRIMESTRAL.xlsx', 'semestral.xlsx') |
    ForEach-Object { Join-Path $outputPath $_ } | Where-Object { Test-Path -LiteralPath $_ -PathType Leaf } |
    ForEach-Object { Get-Item -LiteralPath $_ }
foreach ($file in $finalFiles) {
    if ($file.Length -eq 0) { throw "El archivo final está vacío: $($file.FullName)" }
    Write-Output "ARCHIVO_GENERADO nombre=$($file.Name) bytes=$($file.Length) ruta=$($file.FullName)"
}
Write-Output "TOTAL_GENERADO archivos=$($finalFiles.Count) directorio=$outputPath"
$finalFiles.FullName
