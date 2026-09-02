---
name: generar-mensuales-aios
description: Ejecuta la aplicación Spring Boot plantilla-aios-macro para generar un boletín AIOS mensual consolidado sin correr VBA. Usar cuando el usuario solicite generar, regenerar o procesar un período mensual o cualquier rango inclusivo de meses, incluso entre años.
---

# Generar boletines mensuales AIOS

Ejecutar `scripts/generar-mensuales.ps1` desde PowerShell. El script debe:

1. Verificar que el proyecto contiene `pom.xml`.
2. Iniciar una instancia temporal de la aplicación en un puerto libre, sin reutilizar servicios existentes.
3. Invocar una vez `POST /aios/generar-mensuales` con las fechas inicial y final del rango.
4. Guardar la respuesta como `Boletin_AIOS MENSUAL.xlsx`; el libro debe contener una fila por cada período solicitado, en orden cronológico. Normalizar septiembre como `sep-AA`, aunque la plantilla contenga `sept-AA` fuera de orden.
5. Confirmar que el archivo existe y tiene contenido.
6. Detener al finalizar solamente la instancia que el script haya iniciado, salvo que se solicite conservarla.

## Selección del período

Convertir los períodos solicitados al formato `AAAA-MM` y ejecutar:

- Un solo período: pasar únicamente `-Desde`; el script usa el mismo valor como final.
- Varios períodos: pasar `-Desde` y `-Hasta`; ambos extremos son inclusivos.
- Permitir rangos que crucen años.

Aceptar solicitudes conversacionales como:

```text
Usa Generar mensuales AIOS para generar julio de 2025.
Usa Generar mensuales AIOS para generar junio a diciembre de 2025.
Usa Generar mensuales AIOS para generar noviembre de 2025 a febrero de 2026.
```

Ejecutar directamente el script con:

```powershell
& scripts/generar-mensuales.ps1 -Desde '2025-07'
& scripts/generar-mensuales.ps1 -Desde '2025-06' -Hasta '2025-12'
& scripts/generar-mensuales.ps1 -Desde '2025-11' -Hasta '2026-02'
```

## Parámetros

Usar `-Desde` y opcionalmente `-Hasta` para seleccionar el intervalo. Sin ambos parámetros, conservar junio a diciembre de 2025 como ejecución predeterminada compatible. Usar `-ProjectDir` para otro repositorio, `-OutputDir` para otra salida y `-MavenRepository` para otro repositorio local de dependencias.

## Requisitos y validación

- Requerir conexión y credenciales válidas para Teradata.
- Requerir acceso a los insumos dinámicos de OneDrive que todavía consume `MensualDataReader`; la presentación del boletín proviene de una plantilla interna vacía y no requiere `salidas_referencia`.
- No abrir Excel ni ejecutar macros VBA.
- Fallar ante cualquier respuesta HTTP inválida; no presentar archivos parciales como exitosos.
- Validar el formato de los períodos y rechazar rangos invertidos antes de iniciar la aplicación.
- Entregar al usuario la ruta absoluta del archivo consolidado y cualquier error de generación.
