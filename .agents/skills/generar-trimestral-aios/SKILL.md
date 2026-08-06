---
name: generar-trimestral-aios
description: Actualiza Plantilla AIOS-probable.xlsm ejecutando las rutinas contables de la rama afirmativa de la macro bajar del botón azul y, solo después de guardar y validar la plantilla, ejecuta la aplicación Spring Boot plantilla-aios-macro para generar y verificar el boletín AIOS trimestral. Usar cuando el usuario solicite generar, regenerar o procesar un trimestre con corte al 31 de marzo, 30 de junio, 30 de septiembre o 31 de diciembre.
---

# Generar boletín trimestral AIOS

Ejecutar `scripts/generar-trimestral.ps1` desde PowerShell. Este es el único punto de entrada del flujo.

## Flujo obligatorio

1. Validar que la fecha tenga formato `AAAA-MM-DD`, corresponda al último día de marzo, junio, septiembre o diciembre y que exista `plantillas/Plantilla AIOS-probable.xlsm`.
2. Crear un respaldo fechado de la plantilla.
3. Abrir una instancia aislada y oculta de Excel, escribir la fecha en `CARATULA!B2` y ejecutar `CopiarBalances_BaseMes`, `CopiarBalances_BaseAnual` y `ActualizarSeriesSinPortapapeles`; son las rutinas de actualización que dispara la respuesta **Sí** del botón azul.
4. Esperar la finalización y el recálculo, comprobar el período en `base anual` y validar celdas clave de `cuentas`.
5. Guardar la plantilla únicamente si la actualización termina bien. Ante error o tiempo límite, no continuar con el trimestral.
6. Iniciar una instancia temporal de la aplicación en un puerto libre, sin reutilizar servicios existentes, e invocar `POST /aios/generar?fechaCorte=...&modo=TRIMESTRAL`.
7. Verificar que la salida incluya las ocho hojas esperadas, la etiqueta del período y ningún error común de fórmula.
8. Detener solamente la instancia temporal iniciada por el script, salvo que se solicite conservarla.

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
- Pedir que se cierre `Plantilla AIOS-probable.xlsm` si está abierta; no sobrescribir una copia bloqueada o de solo lectura.
- No presentar un archivo como exitoso si falla la macro, la actualización de la plantilla, la consulta HTTP o la validación final.
- Entregar la ruta absoluta del `.xlsx`, confirmar que la plantilla se actualizó para la misma fecha y comunicar cualquier dependencia que haya impedido completar la macro.
