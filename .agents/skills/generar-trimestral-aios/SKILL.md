---
name: generar-trimestral-aios
description: Genera y verifica uno o varios cortes trimestrales AIOS en un único libro, usando una Plantilla AIOS-probable.xlsm previamente actualizada y contingencia cuando Excel o la macro no pueden completarse. Usar para un corte trimestral o un rango inclusivo de meses que contenga cierres de marzo, junio, septiembre o diciembre.
---

# Generar boletín trimestral AIOS

Para un corte usar `scripts/generar-trimestral.ps1`. Para un rango usar `scripts/generar-trimestrales.ps1`; todos los cierres trimestrales comprendidos deben quedar, en orden cronológico, dentro de un único `Boletin_AIOS TRIMESTRAL.xlsx`.

## Flujo obligatorio

1. Validar que un corte tenga formato `AAAA-MM-DD`, o que los extremos de un rango tengan formato `AAAA-MM`; seleccionar únicamente los cierres de marzo, junio, septiembre y diciembre comprendidos de forma inclusiva. Requerir `plantillas/Plantilla AIOS-probable.xlsm`.
2. Tratar `base anual` y `base mes` como insumos ya actualizados. La skill no debe importar, reemplazar ni modificar estas hojas, ni ejecutar `CopiarBalances_BaseMes` o `CopiarBalances_BaseAnual`.
3. Crear un respaldo fechado de la plantilla.
4. Intentar abrir una instancia aislada y oculta de Excel. Si abre, comprobar que el período solicitado ya exista en `base anual` y `base mes`; ante ausencia del período, detenerse sin guardar ni generar el trimestral. Si la automatización COM no puede iniciar Excel o abrir el libro, registrar la contingencia y continuar con la aplicación usando la plantilla existente.
5. Escribir la fecha en `CARATULA!B2`, ejecutar únicamente `ActualizarSeriesSinPortapapeles`, esperar el recálculo y validar celdas clave de `cuentas`, incluidos valores contables no vacíos y sin ceros anómalos.
6. Guardar la plantilla únicamente si la actualización termina bien. Si la automatización de Excel, la macro, el recálculo o su validación posterior fallan o agotan el tiempo, cerrar Excel sin guardar, registrar `ADVERTENCIA_MACRO_PLANTILLA_OMITIDA` y continuar con la aplicación usando la plantilla existente.
7. Iniciar una instancia temporal de la aplicación en un puerto libre, sin reutilizar servicios existentes. Para un corte invocar `POST /aios/generar`; para un rango invocar una sola vez `POST /aios/generar-rango` con modo `TRIMESTRAL`.
8. Verificar que la salida incluya las ocho hojas esperadas, todas las etiquetas solicitadas en orden cronológico y ningún error común de fórmula; la fila de cada período en `gastos` no debe quedar completamente en cero cuando la fuente contiene información.
9. Detener solamente la instancia temporal iniciada por el script, salvo que se solicite conservarla.

## Ejecución

Usar siempre una fecha de cierre trimestral completa:

```powershell
& .agents/skills/generar-trimestral-aios/scripts/generar-trimestral.ps1 -FechaCorte '2025-09-30'
```

Para un rango inclusivo:

```powershell
& .agents/skills/generar-trimestral-aios/scripts/generar-trimestrales.ps1 -Desde '2025-06' -Hasta '2025-12'
```

El ejemplo produce un solo libro con `jun-25`, `sep-25` y `dic-25`.

Para revisar Excel y los datos previos de la plantilla sin modificar archivos, ejecutar la macro ni generar el boletín:

```powershell
& .agents/skills/generar-trimestral-aios/scripts/generar-trimestral.ps1 -FechaCorte '2025-09-30' -SoloValidar
```

## Parámetros

- Usar `-FechaCorte` para un corte individual. Para un rango usar `-Desde` y opcionalmente `-Hasta` en el script plural; sin `-Hasta`, ambos extremos son el mismo mes.
- Usar `-ProjectDir`, `-PlantillaPath`, `-OutputDir` o `-MavenRepository` solo para reemplazar sus valores predeterminados.
- Usar `-MacroTimeoutMinutes` si la macro necesita más de 30 minutos.
- Usar `-KeepApplicationRunning` únicamente cuando el usuario pida conservar la instancia temporal.

## Requisitos y entrega

- Intentar Microsoft Excel de escritorio para validar y preparar la plantilla. Tratar como recuperable un fallo de automatización al iniciar Excel o abrir el libro, así como un fallo posterior de macro; si Excel sí abre, la ausencia del período y la apertura de solo lectura siguen siendo fatales.
- Requerir que `base anual` y `base mes` ya estén actualizadas dentro de `Plantilla AIOS-probable.xlsm`; la preparación externa de sus fuentes está documentada en `docs/LOGICA_INFORMES_AIOS.md` y queda fuera de la ejecución de esta skill.
- Pedir que se cierre `Plantilla AIOS-probable.xlsm` si está abierta; no sobrescribir una copia bloqueada o de solo lectura.
- Mantener como errores fatales una fecha o plantilla inválida, la ausencia del período en las hojas base, la consulta HTTP, el inicio de la aplicación y la validación final del XLSX.
- Entregar la ruta absoluta del `.xlsx`, los períodos incluidos, el estado de preparación de cada corte, las rutas de los respaldos y, si se omitió una macro, las rutas de sus logs. Un archivo puede considerarse generado cuando Java y la validación final terminan bien, pero nunca afirmar que la plantilla fue actualizada si el estado es `omitida`.
