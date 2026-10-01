#requires -Version 5.1
[CmdletBinding()]
param(
    [string]$Repositorio = (Join-Path $PSScriptRoot '..'),
    [string]$Commit = 'HEAD',
    [Parameter(Mandatory = $true)][string]$Version
)
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
if ($Commit -ne 'HEAD' -and $Commit -notmatch '^[0-9a-fA-F]{7,40}$') {
    throw 'El paquete no contiene un commit Git valido.'
}
if ($Version -notmatch '^(0|[1-9][0-9]*)\.(0|[1-9][0-9]*)$') {
    throw "Version de entrega invalida: $Version"
}
& git -C $Repositorio fetch --no-tags origin '+refs/heads/main:refs/remotes/origin/main'
if ($LASTEXITCODE -ne 0) { throw 'No se pudo consultar main en origin. No se generara ni enviara la entrega.' }
& git -C $Repositorio merge-base --is-ancestor origin/main $Commit
if ($LASTEXITCODE -ne 0) {
    throw "El codigo ($Commit) no contiene el main actual de GitHub. Integre origin/main y genere un paquete nuevo; no cambie solo el numero de version."
}
$pomRemoto = & git -C $Repositorio show origin/main:pom.xml
if ($LASTEXITCODE -ne 0) { throw 'No se pudo leer la version de main.' }
[xml]$pom = $pomRemoto -join "`n"
$versionMain = [string]$pom.project.version
if ($versionMain -notmatch '^\d+\.\d+$') { throw 'main no tiene una version VERSION.RELEASE valida.' }
if ([version]$Version -lt [version]$versionMain) {
    throw "La version $Version es anterior a main ($versionMain). Sincronice pom.xml antes de continuar."
}
Write-Host "Verificado: $Commit contiene origin/main y version $Version >= $versionMain."