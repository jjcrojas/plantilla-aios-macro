---
name: generar-todos-aios
description: "Genera en forma consolidada todos los informes AIOS aplicables a un período o rango inclusivo de meses: un mensual con cada mes, un trimestral con los cierres de marzo, junio, septiembre y diciembre incluidos, y un semestral con los cierres de junio y diciembre incluidos. Usar cuando el usuario solicite generar todos los archivos AIOS, todos los informes, o conjuntamente los boletines mensual, trimestral y semestral para uno o varios períodos."
---

# Generar todos los informes AIOS

Ejecutar `scripts/generar-todos.ps1` como único punto de entrada.

## Períodos

- Convertir cada período solicitado a `AAAA-MM`.
- Para un solo mes, pasar únicamente `-Desde`; se usa también como final.
- Para un rango, pasar `-Desde` y `-Hasta`; ambos extremos son inclusivos.
- Generar siempre un único archivo mensual con todos los meses del rango.
- Generar un único archivo trimestral acumulado si el rango contiene uno o más cierres de marzo, junio, septiembre o diciembre.
- Generar un único archivo semestral acumulado si el rango contiene uno o más cierres de junio o diciembre.
- No crear archivos trimestrales o semestrales vacíos cuando el rango no contenga cortes de esas periodicidades.

Ejemplos:

```powershell
& skills/generar-todos-aios/scripts/generar-todos.ps1 -Desde '2025-06'
& skills/generar-todos-aios/scripts/generar-todos.ps1 -Desde '2025-07' -Hasta '2025-12'
```

El primer ejemplo genera los tres archivos para junio. El segundo genera el mensual de julio a diciembre, el trimestral acumulado con septiembre y diciembre, y el semestral de diciembre.

## Flujo obligatorio

1. Validar el rango antes de iniciar aplicaciones o Excel.
2. Ejecutar la Skill mensual una vez para el intervalo completo.
3. Preparar referencias temporales, sin modificar `salidas_referencia`.
4. Ejecutar los cortes trimestrales en orden cronológico; usar cada salida como referencia del siguiente corte para conservar todas las filas en un solo libro.
5. Ejecutar los cortes semestrales de la misma forma para conservar todas las columnas en un solo libro.
6. Reutilizar los scripts trimestral y semestral, incluidas sus validaciones de Excel, plantilla, fórmulas e insumos.
7. Detener únicamente las instancias temporales iniciadas por el flujo.
8. Entregar las rutas absolutas de los archivos finales y comunicar cualquier corte que no aplique.

## Requisitos

- Requerir conexión y credenciales válidas para Teradata y acceso al servicio web de TRM.
- Requerir Microsoft Excel con macros habilitadas para los cortes trimestrales y semestrales.
- Requerir que `Plantilla AIOS-probable.xlsm` ya contenga los períodos solicitados en `base anual` y `base mes`.
- No presentar como exitoso un conjunto incompleto: ante cualquier error, conservar los registros y reportar exactamente el corte que falló.
- Usar `-SoloPlan` únicamente para validar la selección de períodos sin generar archivos.
