# Versionamiento y generación de entregas

La aplicación usa el mismo esquema numérico de Informes Financieros:
`VERSIÓN.RELEASE`. La base inicial de AIOS es `1.0`.

Para generar una nueva entrega en Windows:

```powershell
.\scripts\crear-paquete-publicacion.ps1
```

El comando incrementa el release (`1.0` → `1.1` → `1.2`). Para un cambio
funcional grande o incompatible:

```powershell
.\scripts\crear-paquete-publicacion.ps1 -CambioMayor
```

Esto incrementa la versión y reinicia el release (`1.2` → `2.0`). La decisión
de cambio mayor es explícita.

El número se persiste en `pom.xml` antes de compilar y no se reutiliza si la
compilación falla. Maven genera `META-INF/build-info.properties` con la versión,
la fecha de compilación y el nombre de la aplicación. La interfaz muestra estos
datos en el pie de página.

Ejecutar Maven directamente recompila la versión registrada en `pom.xml`, pero
no incrementa el número. Para reservar una entrega nueva debe utilizarse el
script. El script no crea etiquetas ni publicaciones de GitHub automáticamente.

El paquete queda en `target\publicacion` e incluye el JAR y un manifiesto con la
versión, la fecha, el commit de Git y el SHA-256 del artefacto.
