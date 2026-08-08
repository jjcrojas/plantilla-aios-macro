---
name: generar-trimestral-aios
description: Valida una Plantilla AIOS-probable.xlsm cuyas hojas base anual y base mes ya fueron actualizadas externamente, actualiza únicamente las series auxiliares y ejecuta la aplicación Spring Boot plantilla-aios-macro para generar y verificar el boletín AIOS trimestral. Usar cuando el usuario solicite generar, regenerar o procesar un trimestre con corte al 31 de marzo, 30 de junio, 30 de septiembre o 31 de diciembre.
---

# Generar boletín trimestral AIOS

Ejecutar `scripts/generar-trimestral.ps1` desde PowerShell. Este es el único punto de entrada del flujo.

## Flujo obligatorio

1. Validar que la fecha tenga formato `AAAA-MM-DD`, corresponda al último día de marzo, junio, septiembre o diciembre y que exista `plantillas/Plantilla AIOS-probable.xlsm`.
2. Tratar `base anual` y `base mes` como insumos ya actualizados. La skill no debe importar, reemplazar ni modificar estas hojas, ni ejecutar `CopiarBalances_BaseMes` o `CopiarBalances_BaseAnual`.
3. Crear un respaldo fechado de la plantilla.
4. Abrir una instancia aislada y oculta de Excel y comprobar que el período solicitado ya exista en `base anual` y `base mes`. Ante ausencia del período, detenerse sin guardar ni generar el trimestral.
5. Escribir la fecha en `CARATULA!B2`, ejecutar únicamente `ActualizarSeriesSinPortapapeles`, esperar el recálculo y validar celdas clave de `cuentas`, incluidos valores contables no vacíos y sin ceros anómalos.
6. Guardar la plantilla únicamente si la actualización termina bien. Ante error o tiempo límite, no continuar con el trimestral.
7. Iniciar una instancia temporal de la aplicación en un puerto libre, sin reutilizar servicios existentes, e invocar `POST /aios/generar?fechaCorte=...&modo=TRIMESTRAL`.
8. Verificar que la salida incluya las ocho hojas esperadas, la etiqueta del período y ningún error común de fórmula; la fila del período en `gastos` no debe quedar completamente en cero cuando la fuente Fox contiene información.
9. Detener solamente la instancia temporal iniciada por el script, salvo que se solicite conservarla.

## Ejecución

Usar siempre una fecha de cierre trimestral completa:

```powershell
& .agents/skills/generar-trimestral-aios/scripts/generar-trimestral.ps1 -FechaCorte '2025-09-30'
```

Para revisar Excel, la plantilla y la macro sin modificar archivos ni generar el boletín:

```powershell
& .agents/skills/generar-trimestral-aios/scripts/generar-trimestral.ps1 -FechaCorte '2025-09-30' -SoloValidar
```

## Parámetros

- Usar `-FechaCorte` obligatoriamente.
- Usar `-ProjectDir`, `-PlantillaPath`, `-OutputDir` o `-MavenRepository` solo para reemplazar sus valores predeterminados.
- Usar `-MacroTimeoutMinutes` si la macro necesita más de 30 minutos.
- Usar `-KeepApplicationRunning` únicamente cuando el usuario pida conservar la instancia temporal.

## Requisitos y entrega

- Requerir Microsoft Excel de escritorio, macros habilitadas, acceso a las rutas configuradas dentro de la macro y credenciales/conexión válidas para Teradata.
- Requerir que `base anual` y `base mes` ya estén actualizadas dentro de `Plantilla AIOS-probable.xlsm`; la preparación externa de sus fuentes está documentada en `docs/LOGICA_INFORMES_AIOS.md` y queda fuera de la ejecución de esta skill.
- Pedir que se cierre `Plantilla AIOS-probable.xlsm` si está abierta; no sobrescribir una copia bloqueada o de solo lectura.
- No presentar un archivo como exitoso si falla la macro, la actualización de la plantilla, la consulta HTTP o la validación final.
- Entregar la ruta absoluta del `.xlsx`, confirmar que la plantilla se actualizó para la misma fecha y comunicar cualquier dependencia que haya impedido completar la macro.
