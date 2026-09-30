[CmdletBinding()]
param(
    [string]$DirectorioSalida,
    [switch]$OmitirPruebas,
    [switch]$ExigirGitLimpio,
    [switch]$CambioMayor
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$raizProyecto = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..')).Path
$pomPath = Join-Path $raizProyecto 'pom.xml'

if (-not (Get-Command mvn -ErrorAction SilentlyContinue)) {
    throw 'Maven (mvn) no está disponible en PATH.'
}
if (-not (Get-Command git -ErrorAction SilentlyContinue)) {
    throw 'Git no está disponible en PATH.'
}

if ([string]::IsNullOrWhiteSpace($DirectorioSalida)) {
    $DirectorioSalida = Join-Path $raizProyecto 'target\publicacion'
} elseif (-not [IO.Path]::IsPathRooted($DirectorioSalida)) {
    $DirectorioSalida = Join-Path $raizProyecto $DirectorioSalida
}
$DirectorioSalida = [IO.Path]::GetFullPath($DirectorioSalida)

$estadoGit = @(& git -C $raizProyecto status --porcelain)
if ($LASTEXITCODE -ne 0) {
    throw 'No fue posible consultar el estado de Git.'
}
if ($estadoGit.Count -gt 0 -and $ExigirGitLimpio) {
    throw 'El repositorio tiene cambios sin confirmar.'
}
if ($estadoGit.Count -gt 0) {
    Write-Warning 'El paquete incluirá cambios locales sin confirmar.'
}

[xml]$pom = Get-Content -LiteralPath $pomPath -Raw
$version = [string]$pom.project.version
if ($version -notmatch '^(0|[1-9][0-9]*)\.(0|[1-9][0-9]*)$') {
    throw "La versión debe tener formato numérico VERSION.RELEASE: $version"
}

$numeroVersion = [int]$Matches[1]
$release = [int]$Matches[2]
$versionNueva = if ($CambioMayor) { "$($numeroVersion + 1).0" } else { "$numeroVersion.$($release + 1)" }
$contenidoPom = [IO.File]::ReadAllText($pomPath)
$patronVersion = '<version>' + [regex]::Escape($version) + '</version>'
$contenidoPom = [regex]::Replace($contenidoPom, $patronVersion, "<version>$versionNueva</version>", 1)
[IO.File]::WriteAllText($pomPath, $contenidoPom, [Text.UTF8Encoding]::new($false))
$tieneCambios = $true

$repositorioMaven = Join-Path $env:USERPROFILE '.m2\repository'
$argumentosMaven = @('-f', $pomPath, "-Dmaven.repo.local=$repositorioMaven", 'clean', 'package')
if ($OmitirPruebas) {
    $argumentosMaven = @('-f', $pomPath, "-Dmaven.repo.local=$repositorioMaven", '-DskipTests', 'clean', 'package')
}

Write-Host "Compilando versión $versionNueva..."
& mvn @argumentosMaven
if ($LASTEXITCODE -ne 0) {
    throw "Maven terminó con código $LASTEXITCODE. El número $versionNueva quedó reservado."
}

$jar = Get-ChildItem -LiteralPath (Join-Path $raizProyecto 'target') -Filter '*.jar' -File |
    Where-Object { $_.Name -notlike '*.original' } |
    Sort-Object LastWriteTimeUtc -Descending |
    Select-Object -First 1
if (-not $jar) {
    throw 'La compilación terminó, pero no se encontró el JAR ejecutable.'
}

$commit = (& git -C $raizProyecto rev-parse --short HEAD).Trim()
$hashJar = (Get-FileHash -LiteralPath $jar.FullName -Algorithm SHA256).Hash.ToLowerInvariant()
$marcaTiempo = Get-Date -Format 'yyyyMMdd-HHmmss'
$directorioTemporal = Join-Path ([IO.Path]::GetTempPath()) ("aios-publicacion-" + [guid]::NewGuid().ToString('N'))
$zipPath = Join-Path $DirectorioSalida "plantilla-aios-$versionNueva-$marcaTiempo.zip"

try {
    New-Item -ItemType Directory -Path $DirectorioSalida -Force | Out-Null
    New-Item -ItemType Directory -Path $directorioTemporal -Force | Out-Null
    Copy-Item -LiteralPath $jar.FullName -Destination (Join-Path $directorioTemporal $jar.Name)

    $manifiesto = @(
        "app.version=$versionNueva",
        "package.created-at=$((Get-Date).ToUniversalTime().ToString('o'))",
        "git.commit=$commit",
        "git.dirty=$($tieneCambios.ToString().ToLowerInvariant())",
        "jar.name=$($jar.Name)",
        "jar.sha256=$hashJar"
    ) -join "`n"
    [IO.File]::WriteAllText(
        (Join-Path $directorioTemporal 'publicacion-manifest.properties'),
        $manifiesto + "`n",
        [Text.UTF8Encoding]::new($false)
    )

    Compress-Archive -Path (Join-Path $directorioTemporal '*') -DestinationPath $zipPath -Force
    Write-Host "Paquete generado: $zipPath" -ForegroundColor Green
    Write-Host "JAR: $($jar.Name)"
    Write-Host "SHA256: $hashJar"
} finally {
    $tempBase = [IO.Path]::GetFullPath([IO.Path]::GetTempPath())
    $tempCompleto = [IO.Path]::GetFullPath($directorioTemporal)
    if ((Test-Path -LiteralPath $directorioTemporal) -and
        $tempCompleto.StartsWith($tempBase, [StringComparison]::OrdinalIgnoreCase)) {
        Remove-Item -LiteralPath $directorioTemporal -Recurse -Force
    }
}
