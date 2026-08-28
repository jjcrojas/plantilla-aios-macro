---
name: generar-semestral-aios
description: Genera y verifica uno o varios cortes semestrales AIOS en un único libro, intentando preparar la Plantilla AIOS-probable.xlsm y usando contingencia cuando Excel o la macro no pueden completarse. Usar para un corte al 30 de junio o 31 de diciembre, o para un rango inclusivo que contenga esos cierres.
---

# Generar semestral AIOS

Para un corte usar `scripts/generar-semestral.ps1`. Para un rango usar `scripts/generar-semestrales.ps1`; todos los cierres semestrales comprendidos deben quedar, en orden cronológico, como columnas de un único `semestral.xlsx`.

## Flujo obligatorio

1. Validar que un corte tenga formato `AAAA-MM-DD`, o que los extremos de un rango tengan formato `AAAA-MM`; seleccionar de forma inclusiva únicamente el 30 de junio y el 31 de diciembre.
2. Comprobar que exista `plantillas/Plantilla AIOS-probable.xlsm`. Si Excel puede abrirla, comprobar que las hojas `base anual` y `base mes` ya contengan el período solicitado; la skill no reemplaza ni modifica estas dos hojas. Si la automatización COM no puede iniciar Excel o abrir el libro, registrar la contingencia y continuar con la aplicación usando la plantilla existente.
3. Crear un respaldo fechado de la plantilla.
4. Abrir una instancia aislada y oculta de Excel, escribir la fecha en `CARATULA!B2` e intentar ejecutar la macro no interactiva `ActualizarSeriesSinPortapapeles`.
5. Esperar el recálculo, validar las celdas de `cuentas` que consume el semestral y guardar la plantilla solo si la preparación termina correctamente. Si la automatización de Excel, la macro, el recálculo o su validación posterior fallan o agotan el tiempo, cerrar Excel sin guardar, registrar `ADVERTENCIA_MACRO_PLANTILLA_OMITIDA` y continuar con la aplicación usando la plantilla existente.
6. Iniciar una instancia temporal de la aplicación en un puerto libre. Para un corte invocar `POST /aios/generar`; para un rango invocar una sola vez `POST /aios/generar-rango` con modo `SEMESTRAL`.
7. Verificar que la salida sea un libro XLSX válido, incluya en orden cronológico una columna por cada semestre solicitado, contenga valores numéricos y no tenga errores comunes de fórmula.
8. Detener solamente la instancia temporal iniciada por el script, salvo que se solicite conservarla.

No ejecutar la macro heredada `bajar`: es interactiva, abre insumos adicionales e intenta producir una salida semestral por VBA. La preparación requerida antes de Java es `ActualizarSeriesSinPortapapeles`.

## Ejecución

```powershell
& '.agents\skills\generar-semestral-aios\scripts\generar-semestral.ps1' -FechaCorte '2025-06-30'
```

Para un rango inclusivo:

```powershell
& '.agents\skills\generar-semestral-aios\scripts\generar-semestrales.ps1' -Desde '2025-06' -Hasta '2025-12'
```

El ejemplo produce un solo libro con las columnas junio de 2025 y diciembre de 2025.

Para validar dependencias y datos previos de la plantilla sin ejecutar la macro, guardar cambios ni generar el archivo:

```powershell
& '.agents\skills\generar-semestral-aios\scripts\generar-semestral.ps1' -FechaCorte '2025-06-30' -SoloValidar
```

## Resultado esperado

Entregar la ruta absoluta de `semestral.xlsx`, los períodos validados, el tamaño, el estado de preparación de cada corte y las rutas de los respaldos. Un rango se escribe por defecto en `target/aios-output/semestrales-AAAA-MM-a-AAAA-MM/semestral.xlsx`.

Si falla la automatización de Excel al iniciar o abrir el libro, o falla la macro posteriormente, informar la advertencia y las rutas de los logs, continuar y aceptar el semestral solo si Java y la validación final terminan bien. La fecha o plantilla inválida, la ausencia del período cuando Excel logra validarlo, la apertura de solo lectura, el fallo de la aplicación o un libro que no supera las validaciones siguen siendo errores fatales. No afirmar que la plantilla fue actualizada cuando el estado sea `omitida`.
