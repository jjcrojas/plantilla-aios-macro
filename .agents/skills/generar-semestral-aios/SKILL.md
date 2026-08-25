---
name: generar-semestral-aios
description: Valida la Plantilla AIOS-probable.xlsm para un corte semestral, intenta ejecutar su rutina no interactiva de actualización y después ejecuta la aplicación Spring Boot plantilla-aios-macro para generar y verificar semestral.xlsx, incluso cuando la macro no puede completarse. Usar cuando el usuario solicite generar, regenerar o procesar un informe semestral con corte al 30 de junio o 31 de diciembre.
---

# Generar semestral AIOS

Ejecutar `scripts/generar-semestral.ps1` desde PowerShell. Este es el único punto de entrada del flujo.

## Flujo obligatorio

1. Validar que la fecha tenga formato `AAAA-MM-DD` y corresponda al 30 de junio o al 31 de diciembre.
2. Comprobar que exista `plantillas/Plantilla AIOS-probable.xlsm` y que las hojas `base anual` y `base mes` ya contengan el período solicitado. La skill no reemplaza ni modifica estas dos hojas.
3. Crear un respaldo fechado de la plantilla.
4. Abrir una instancia aislada y oculta de Excel, escribir la fecha en `CARATULA!B2` e intentar ejecutar la macro no interactiva `ActualizarSeriesSinPortapapeles`.
5. Esperar el recálculo, validar las celdas de `cuentas` que consume el semestral y guardar la plantilla solo si la preparación termina correctamente. Si la macro, el recálculo o su validación posterior fallan o agotan el tiempo después de superar las validaciones previas, cerrar Excel sin guardar, registrar `ADVERTENCIA_MACRO_PLANTILLA_OMITIDA` y continuar con la aplicación usando la plantilla existente.
6. Iniciar una instancia temporal de la aplicación en un puerto libre e invocar `POST /aios/generar?fechaCorte=...&modo=SEMESTRAL`.
7. Verificar que la salida sea un libro XLSX válido, incluya la columna del semestre solicitado, contenga valores numéricos y no tenga errores comunes de fórmula.
8. Detener solamente la instancia temporal iniciada por el script, salvo que se solicite conservarla.

No ejecutar la macro heredada `bajar`: es interactiva, abre insumos adicionales e intenta producir una salida semestral por VBA. La preparación requerida antes de Java es `ActualizarSeriesSinPortapapeles`.

## Ejecución

```powershell
& '.agents\skills\generar-semestral-aios\scripts\generar-semestral.ps1' -FechaCorte '2025-06-30'
```

Para validar dependencias y datos previos de la plantilla sin ejecutar la macro, guardar cambios ni generar el archivo:

```powershell
& '.agents\skills\generar-semestral-aios\scripts\generar-semestral.ps1' -FechaCorte '2025-06-30' -SoloValidar
```

## Resultado esperado

Entregar la ruta absoluta de `semestral.xlsx`, el período validado, el tamaño, el estado `preparacionPlantilla=actualizada|omitida` y la ruta del respaldo. Por defecto se escribe en `target/aios-output/semestral-AAAA-MM/semestral.xlsx`.

Si falla la macro después de superar las validaciones previas, informar la advertencia y las rutas de los logs, continuar y aceptar el semestral solo si Java y la validación final terminan bien. La fecha o plantilla inválida, la ausencia del período en las hojas base, la apertura de solo lectura, el fallo de la aplicación o un libro que no supera las validaciones siguen siendo errores fatales. No afirmar que la plantilla fue actualizada cuando el estado sea `omitida`.
