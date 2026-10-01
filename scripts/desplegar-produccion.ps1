#requires -Version 5.1
[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [string]$Paquete,
    [string]$Destino = '\\172.19.130.163\Temp',
    [string]$RutaWslDestino = '/mnt/c/Temp',
    [switch]$OmitirPruebas,
    [switch]$ExigirGitLimpio,
    [switch]$CambioMayor
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
$paqueteExistente = -not [string]::IsNullOrWhiteSpace($Paquete)
$raiz = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..')).Path
if ($Paquete -and ($OmitirPruebas -or $ExigirGitLimpio -or $CambioMayor)) {
    throw 'Con -Paquete no se compila ni se admiten opciones de compilacion.'
}
if ($RutaWslDestino -notmatch '^/mnt/[a-z]/' -or $RutaWslDestino -match "['`r`n]") {
    throw 'RutaWslDestino debe ser una ruta /mnt/<unidad>/carpeta sin comillas simples ni saltos de linea.'
}
if (-not (Test-Path -LiteralPath $Destino -PathType Container)) {
    throw "No se puede acceder a $Destino. Compruebe VPN/red, recurso compartido y permisos SMB."
}
if (-not $PSCmdlet.ShouldProcess($Destino, 'Crear o validar el paquete y transferir una entrega verificada (no ejecuta WSL)')) {
    return
}

if (-not $Paquete) {
    $salida = Join-Path $raiz ('target\envios\' + [guid]::NewGuid().ToString('N'))
    & (Join-Path $PSScriptRoot 'crear-paquete-publicacion.ps1') -DirectorioSalida $salida `
        -OmitirPruebas:$OmitirPruebas -ExigirGitLimpio:$ExigirGitLimpio -CambioMayor:$CambioMayor
    $archivos = @(Get-ChildItem -LiteralPath $salida -Filter '*.zip' -File)
    if ($archivos.Count -ne 1) { throw 'El empaquetador no produjo exactamente un ZIP.' }
    $Paquete = $archivos[0].FullName
}
$archivo = Get-Item -LiteralPath $Paquete -ErrorAction Stop
if ($archivo.PSIsContainer -or $archivo.Extension -ne '.zip') { throw 'Paquete debe ser un archivo ZIP.' }
if ($archivo.Name -match "['`r`n]") { throw 'El nombre del ZIP no puede contener comillas simples ni saltos de linea.' }

# Validar el JAR y usar el publicador de la misma entrega, no el de otra version.
Add-Type -AssemblyName System.IO.Compression.FileSystem
$zip = [IO.Compression.ZipFile]::OpenRead($archivo.FullName)
try {
    $manifest = $zip.GetEntry('publicacion-manifest.properties')
    $publicador = $zip.GetEntry('scripts/publicar-produccion.sh')
    $gestor = $zip.GetEntry('scripts/manage-app.sh')
    $validador = $zip.GetEntry('scripts/validar-paquete.py')
    if (-not $manifest -or -not $publicador -or -not $gestor -or -not $validador) {
        throw 'El ZIP no es un paquete de publicacion de AIOS.'
    }
    $reader = [IO.StreamReader]::new($manifest.Open())
    try { $contenido = $reader.ReadToEnd() } finally { $reader.Dispose() }
    if ($paqueteExistente) {
        $commitPaquete = [regex]::Match($contenido, '(?m)^git.commit=([^\r\n]+)').Groups[1].Value
        $versionPaquete = [regex]::Match($contenido, '(?m)^app.version=([^\r\n]+)').Groups[1].Value
        & (Join-Path $PSScriptRoot 'verificar-base-publicacion.ps1') -Repositorio $raiz -Commit $commitPaquete -Version $versionPaquete
    }
    $nombreJar = [regex]::Match($contenido, '(?m)^jar.name=([^\r\n]+)').Groups[1].Value
    $hashEsperado = [regex]::Match($contenido, '(?m)^jar.sha256=([a-fA-F0-9]{64})\r?$').Groups[1].Value
    if ($nombreJar -notmatch '^[^/\\]+\.jar$' -or -not $hashEsperado) { throw 'Manifiesto de JAR invalido.' }
    $jar = $zip.GetEntry('' + $nombreJar)
    if (-not $jar) { throw 'Falta el JAR indicado en el manifiesto.' }
    $stream = $jar.Open()
    $sha = [Security.Cryptography.SHA256]::Create()
    try { $hashJar = [BitConverter]::ToString($sha.ComputeHash($stream)).Replace('-', '') }
    finally { $sha.Dispose(); $stream.Dispose() }
    if ($hashJar -ne $hashEsperado) { throw 'El SHA256 del JAR no coincide con el manifiesto.' }
    $reader = [IO.StreamReader]::new($validador.Open())
    try { $textoValidador = $reader.ReadToEnd().Replace([string][char]13, '') } finally { $reader.Dispose() }
    $reader = [IO.StreamReader]::new($publicador.Open())
    try { $textoPublicador = $reader.ReadToEnd().Replace("`r`n", "`n") } finally { $reader.Dispose() }
} finally { $zip.Dispose() }

$nombreEntrega = 'aios-' + (Get-Date -Format 'yyyyMMdd-HHmmss') + '-' + [guid]::NewGuid().ToString('N').Substring(0, 8)
$pendiente = Join-Path $Destino ($nombreEntrega + '.partial')
$entrega = Join-Path $Destino $nombreEntrega
$temporalValidador = Join-Path ([IO.Path]::GetTempPath()) ([guid]::NewGuid().ToString('N') + '.py')
$temporal = Join-Path ([IO.Path]::GetTempPath()) ([guid]::NewGuid().ToString('N') + '.sh')
try {
    [IO.File]::WriteAllText($temporal, $textoPublicador, [Text.UTF8Encoding]::new($false))
    [IO.File]::WriteAllText($temporalValidador, $textoValidador, [Text.UTF8Encoding]::new($false))
    New-Item -ItemType Directory -Path $pendiente -ErrorAction Stop | Out-Null
    foreach ($item in @(
        @{ Origen = $archivo.FullName; Nombre = $archivo.Name },
        @{ Origen = $temporal; Nombre = 'publicar-produccion.sh' },
        @{ Origen = $temporalValidador; Nombre = 'validar-paquete.py' }
    )) {
        $hashOrigen = (Get-FileHash -LiteralPath $item.Origen -Algorithm SHA256).Hash
        $rutaRemota = Join-Path $pendiente $item.Nombre
        Copy-Item -LiteralPath $item.Origen -Destination $rutaRemota -ErrorAction Stop
        $hashDestino = (Get-FileHash -LiteralPath $rutaRemota -Algorithm SHA256).Hash
        if ($hashOrigen -ne $hashDestino) { throw "Fallo de integridad al transferir $($item.Nombre)." }
    }
    # La carpeta solo se publica con su nombre definitivo despues de verificar ambos archivos.
    Rename-Item -LiteralPath $pendiente -NewName $nombreEntrega -ErrorAction Stop
} catch {
    throw "Transferencia incompleta. No ejecute archivos de '$pendiente'. $($_.Exception.Message)"
} finally {
    if (Test-Path -LiteralPath $temporalValidador) { Remove-Item -LiteralPath $temporalValidador -Force }
    if (Test-Path -LiteralPath $temporal) { Remove-Item -LiteralPath $temporal -Force }
}

$rutaWsl = $RutaWslDestino.TrimEnd('/') + '/' + $nombreEntrega
$comando = "bash '$rutaWsl/publicar-produccion.sh' '$rutaWsl/$($archivo.Name)'"
Write-Host "Transferencia verificada: $entrega" -ForegroundColor Green
Write-Host 'La aplicacion aun no se ha instalado ni reiniciado. Ejecute en WSL de produccion:'
Write-Host $comando
[pscustomobject]@{ DirectorioRemoto = $entrega; Paquete = $archivo.Name; ComandoWsl = $comando; Estado = 'Transferido y verificado; instalacion pendiente' }