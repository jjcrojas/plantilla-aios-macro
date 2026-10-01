#!/usr/bin/env python3
"""Valida y extrae exclusivamente el contenido de una entrega AIOS."""
import hashlib
import pathlib
import re
import sys
import zipfile

archive, destination = sys.argv[1:]
root = pathlib.Path(destination)
with zipfile.ZipFile(archive) as z:
    names = z.namelist()
    if len(names) != len(set(names)):
        raise ValueError('ZIP con entradas duplicadas')
    props = dict(line.split('=', 1) for line in z.read('publicacion-manifest.properties').decode('utf-8').splitlines() if '=' in line)
    jar = props.get('jar.name', '')
    if not re.fullmatch(r'[A-Za-z0-9_.-]+\.jar', jar):
        raise ValueError('Nombre de JAR invalido')
    if not re.fullmatch(r'\d+\.\d+', props.get('app.version', '')):
        raise ValueError('Version invalida')
    required = {jar, 'publicacion-manifest.properties', '.env.example', 'scripts/manage-app.sh', 'scripts/publicar-produccion.sh', 'scripts/validar-paquete.py'}
    allowed = required | {'scripts/desplegar-produccion.ps1', 'scripts/verificar-base-publicacion.ps1', 'docs/publicacion-produccion.md'}
    for info in z.infolist():
        if info.is_dir():
            if info.filename not in ('scripts/', 'docs/'):
                raise ValueError('Directorio inesperado')
        elif info.filename not in allowed or (info.external_attr >> 16) & 0o170000 == 0o120000:
            raise ValueError('Entrada no permitida: ' + info.filename)
    if not required.issubset(names):
        raise ValueError('Paquete incompleto')
    h = hashlib.sha256()
    with z.open(jar) as f:
        for block in iter(lambda: f.read(1024 * 1024), b''):
            h.update(block)
    if h.hexdigest() != props.get('jar.sha256'):
        raise ValueError('Hash del JAR incorrecto')
    z.extractall(root)
    print(props['app.version'])
