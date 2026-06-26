# Fórmulas y fuentes de datos de informes AIOS

## 1. Convenciones

| Convención | Significado |
|---|---|
| `TRM` | Tasa representativa del mercado leída de `PIB_PEA_TRM_DG`; si llega en cero se usa `1` para evitar división por cero. |
| `pct(x)` | `x * 100`, usado cuando la plantilla espera puntos porcentuales. |
| `safeDivide(a,b)` | `a / b` con escala 8 y redondeo `HALF_UP`; retorna cero si el denominador es nulo o cero. |
| `MM` | Millones. |
| `COP` | Pesos colombianos. |
| `USD` | Dólares. |
| `Valor crudo` | Conteo o número sin conversión monetaria. |

## 2. Informe mensual (`Boletin_AIOS MENSUAL.xlsx`, hoja `HOJA1`)

| Columna destino | Fórmula / dato | Fuente, hoja y celda | Unidad escrita |
|---:|---|---|---|
| B | `afiliados` | Query Teradata sobre `PROD_DWH_CONSULTA.FORMATO491` (`RENGLON=999`, suma `TOTAL_AFILIADOS_TOTAL`, fondos 1000/5000/6000/7000/8000) | Valor crudo (personas) |
| C | `aportantes` | Query Teradata sobre `PROD_DWH_CONSULTA.FORMATO491` (`RENGLON=999`, suma `TOTAL_AFILIADOS_COTIZANTES`, fondos 1000/5000/6000/7000/8000, sin filtro por `CODIGO_ENTIDAD`) | Valor crudo (personas) |
| D | `traspasosSistema` | Formato 493, `Traslados Entre AFP`, `BQ11` | Valor crudo |
| E | `vrFondo / TRM` | `SISTEMA TOTAL`, hoja `restot`, total sistema; TRM de `PIB_PEA_TRM_DG` | USD |
| F | `total1 / TRM` | `LIMITES`, hoja `AIOS`, `AB4`; TRM de `PIB_PEA_TRM_DG` | USD |
| G | `dudaG * 100` | `LIMITES`, hoja `AIOS`, `C4` | Porcentaje |
| H | `dudaEf * 100` | `LIMITES`, hoja `AIOS`, `E4` | Porcentaje |
| I | `dudaNf * 100` | `LIMITES`, hoja `AIOS`, `G4` | Porcentaje |
| J | `dudaAc * 100` | `LIMITES`, hoja `AIOS`, `I4` | Porcentaje |
| K | `dudaF * 100` | `LIMITES`, hoja `AIOS`, `K4` | Porcentaje |
| L | `h17 * 100` | `LIMITES`, hoja `AIOS`, rango `O4:Y4` | Porcentaje |
| M | `otros * 100` | `LIMITES`, hoja `AIOS`, `AA4` | Porcentaje |
| N | `tmpNominal1 * 100` | `Rent_Vr_Uni_Moderado`, primera hoja, `D11` | Porcentaje |
| O | `tmpReal1 * 100` | `Rent_Vr_Uni_Moderado`, primera hoja, `D10` | Porcentaje |
| P | `4` | Constante | Valor crudo |
| Q | `consFdosAdmon` | Query Teradata Formato 491: concentración de afiliados/personas = dos AFP con más `TOTAL_AFILIADOS_TOTAL` / total sistema | Porcentaje |
| R | `porcVrFondo` | `SISTEMA TOTAL`, `restot`, `(Protección + Porvenir) / Sistema` | Ratio / porcentaje según plantilla |
| S | `TRM` | `PIB_PEA_TRM_DG`, último valor aplicable a la fecha de corte | COP por USD |

## 3. Informe trimestral

El informe trimestral escribe mapas por hoja. Las fórmulas exactas dependen de los mapas generados por `TrimestralDataReader`; la tabla resume la forma de escritura y unidad.

| Hoja | Fórmula / mapeo | Fuente principal | Unidad escrita |
|---|---|---|---|
| `afiliados` | Valores por fondo y administradora (`mod_*`, `con_*`, `mr_*`, combinaciones) | Formato 491 / archivos trimestrales de referencia | Valor crudo (personas) |
| `aportantes` | `colf`, `porv`, `prot`, `sk` y ceros para entidades no aplicables | Query Teradata Formato 491 filtrada por `CODIGO_ENTIDAD` para cada AFP: 10, 3, 2 y 9 | Valor crudo (personas) |
| `colombia` | Saldos por fondo/administradora, con agregados como `mod_sk + mod_alt` | `SISTEMA TOTAL` y datos de fondos | USD o MM USD según plantilla |
| `traspasos` | Traspasos por administradora | Formato 493 | Valor crudo |
| `gastos` | `gastoNetoCOP / TRM` | Base anual / cuentas de gasto; TRM de `PIB_PEA_TRM_DG` | USD |
| `promotores` | Ceros cuando no hay fuente disponible | Constante temporal | Valor crudo |
| `rentabilidad` | Rentabilidad nominal y real por administradora/fondo | `Rent_Vr_Uni_Moderado` y/o series de rentabilidad | Porcentaje |
| `comisiones` | Comisiones obligatorias por administradora (`col_obl`, `por_obl`, `pro_obl`, `ska_obl`) | Datos trimestrales de comisiones | Porcentaje |

## 4. Informe semestral (`semestral.xlsx`)

| Fila | Fórmula / dato | Fuente, hoja y celda | Unidad escrita |
|---:|---|---|---|
| 3 | `afiliados activos` | Query Teradata sobre `PROD_DWH_CONSULTA.FORMATO491` (`RENGLON=999`, suma `TOTAL_AFILIADOS_ACTIVOS_TOTAL`, fondos 1000/5000/6000/7000/8000) | Valor crudo (personas) |
| 4 | `(afiliadosMenor30 / afiliados) * 100` | Query Teradata Formato 491 con reglas de subcuenta/unidad de captura; denominador = total afiliados (`RENGLON=999`) | Porcentaje |
| 5 | `(afiliados30a44 / afiliados) * 100` | Query Teradata Formato 491 con reglas de subcuenta/unidad de captura; denominador = total afiliados (`RENGLON=999`) | Porcentaje |
| 6 | `(afiliados45a59 / afiliados) * 100` | Query Teradata Formato 491 con reglas de subcuenta/unidad de captura; denominador = total afiliados (`RENGLON=999`) | Porcentaje |
| 7 | `(afiliadosMayor60 / afiliados) * 100` | Query Teradata Formato 491 con reglas de subcuenta/unidad de captura; denominador = total afiliados (`RENGLON=999`) | Porcentaje |
| 8 | `100` | Constante | Porcentaje |
| 9 | `afiliados / 1000` | Afiliados totales por query Teradata Formato 491 (`TOTAL_AFILIADOS_TOTAL`, `RENGLON=999`) | Miles de personas |
| 10 | `(mujeres / afiliados) * 100` | Query Teradata Formato 491 (`SUM(TOTAL_AFILIADOS_M)`, `RENGLON=999`, fondos 1000/5000/6000/7000/8000); denominador = total afiliados por query | Porcentaje |
| 11 | `aportantes` | Query Teradata sobre Formato 491 (`SUM(TOTAL_AFILIADOS_COTIZANTES)`, `RENGLON=999`, fondos 1000/5000/6000/7000/8000, sin filtro por `CODIGO_ENTIDAD`) | Valor crudo (personas) |
| 12 | `(afiliados / PEA) * 100` | Afiliados totales por query Teradata Formato 491; PEA de `PIB_PEA_TRM_DG` | Porcentaje |
| 13 | `(aportantes / PEA) * 100` | Aportantes por query Teradata del Formato 491; PEA de `PIB_PEA_TRM_DG` | Porcentaje |
| 14 | `(aportantes / afiliados) * 100` | Aportantes y afiliados por query Teradata del Formato 491 | Porcentaje |
| 15 | `salario mínimo ponderado COP / TRM` | Query Teradata Formato 491 para IBC ponderado con salario oficial de `SalarioMinimo.csv`; TRM de `PIB_PEA_TRM_DG` | USD |
| 16 | `total pensionados` | Formato 495, `TOTAL PENSIONADOS`, parámetro `B4`, valor en columna `I` para la fecha | Valor crudo (personas) |
| 17 | `por Entidad!BI62 / fila16` | Formato 495, hoja `por Entidad`, parámetro `C6`, celda `BI62` | Ratio / porcentaje |
| 18 | `por Entidad!BH62 / fila16` | Formato 495, hoja `por Entidad`, parámetro `C6`, celda `BH62` | Ratio / porcentaje |
| 19 | `por Entidad!BJ62 / fila16` | Formato 495, hoja `por Entidad`, parámetro `C6`, celda `BJ62` | Ratio / porcentaje |
| 25 | `Formato 493!M11 / 1000` | `Serie_Formato_493 MOVIMIENTO AFILIADOS.xlsx`, hoja `Fallecidos`; se escribe la fecha de corte en `B11`, `D4=99`, y se toma `M11` | Miles |
| 26 | `traspasosSistema` | Formato 493 / lector mensual | Valor crudo |
| 27 | `traspasosSistema / afiliados` | Traspasos de Formato 493 y afiliados totales por query Teradata Formato 491 | Porcentaje (formato Excel) |
| 28 | `(fondoSistemaJ14 * 1000 / TRM) / 1,000,000` | `SISTEMA TOTAL`, `restot`, `J14`; TRM de `PIB_PEA_TRM_DG` | MM USD |
| 29 | `fila28 / (pibSemestral / TRM)` | PIB semestral y TRM de `PIB_PEA_TRM_DG` | Ratio / porcentaje |
| 30 | `total1 / TRM` | `LIMITES`, hoja `AIOS`, total equivalente; TRM | USD |
| 31 | `dudaG` | `LIMITES`, hoja `AIOS`, `C4` | Ratio |
| 32 | `dudaEf` | `LIMITES`, hoja `AIOS`, `E4` | Ratio |
| 33 | `dudaNf` | `LIMITES`, hoja `AIOS`, `G4` | Ratio |
| 34 | `dudaAc` | `LIMITES`, hoja `AIOS`, `I4` | Ratio |
| 35 | `dudaF` | `LIMITES`, hoja `AIOS`, `K4` | Ratio |
| 36 | `0` | Constante | Valor crudo |
| 37 | `dudaGe` | `LIMITES`, hoja `AIOS` | Ratio |
| 38 | `dudaEfe` | `LIMITES`, hoja `AIOS` | Ratio |
| 39 | `dudaNfe` | `LIMITES`, hoja `AIOS` | Ratio |
| 40 | `dudaAce` | `LIMITES`, hoja `AIOS` | Ratio |
| 41 | `dudaFe` | `LIMITES`, hoja `AIOS` | Ratio |
| 42 | `2` | Constante | Valor crudo |
| 43 | `otros` | `LIMITES` / lector mensual | Ratio |
| 44 | `(O4 + Q4 + S4 + U4 + W4 + Y4) * 100` | `LIMITES`, hoja `AIOS`, celdas indicadas | Porcentaje |
| 45 | `fila28 / deudaGubernamentalTotalUSD` | `PIB_PEA_TRM_DG`, hoja `Hoja1`, fecha en columna `L`, deuda en columna `M` | Porcentaje (formato Excel) |
| 46 | `4` | Constante | Valor crudo |
| 47 | `(restot!C14 + restot!D14) / restot!J14` | `SISTEMA TOTAL`, hoja `restot`, `C14`, `D14`, `J14` | Porcentaje (formato Excel) |
| 48 | `activosCuentas / TRM` | `Plantilla AIOS-probable`, hoja `CUENTAS`, `C6`; TRM | USD |
| 49 | `pasivosCuentas / TRM` | `Plantilla AIOS-probable`, hoja `CUENTAS`, `C4`; TRM | USD |
| 50 | `(activosCuentas - pasivosCuentas) / TRM` | `Plantilla AIOS-probable`, hoja `CUENTAS`, `C6` y `C4`; TRM | USD |
| 51 | `comisiones` | `Plantilla AIOS-probable`, hoja `CUENTAS`, `E13` | COP / unidad de plantilla |
| 52 | `gastos` | `Plantilla AIOS-probable`, hoja `CUENTAS`, `G15` | COP / unidad de plantilla |
| 53 | `resultadoOperacion` | `Plantilla AIOS-probable`, hoja `CUENTAS`, `E41` | COP / unidad de plantilla |
| 54 | `resultadoNeto` | `Plantilla AIOS-probable`, hoja `CUENTAS`, `E44` | COP / unidad de plantilla |
| 55 | `admon` | `Plantilla AIOS-probable`, hoja `CUENTAS`, `H24` | COP / unidad de plantilla |
| 56 | `C21 / TRM` | `Plantilla AIOS-probable`, hoja `cuentas`, cuenta `511500`, celda `C21`; TRM | USD |
| 57 | `C22 / TRM` | `Plantilla AIOS-probable`, hoja `cuentas`, cuenta `511527`, celda `C22`; TRM | USD |
| 58 | `(C21 + C22) / TRM` | `Plantilla AIOS-probable`, hoja `cuentas`, cuentas `511500` y `511527`; TRM | USD |
| 59 | `(C24 + C28 + C29 + C31 + C32 + C33 + C34 + C35 + C36 + C37 + C38) / TRM` | `Plantilla AIOS-probable`, hoja `cuentas`, cuentas `512000`, `513000`, `513500`, `514000`, `514500`, `515000`, `515500`, `516000`, `516500`, `517000`, `517200`; TRM | USD |
| 60 | `C15 / TRM` | `Plantilla AIOS-probable`, hoja `cuentas`, cuenta `510000`, celda `C15`; TRM | USD |
| 61 | `(aportesRecibidos136 / TRM) / (aportantes / 1000) * 1000` | `Formato_136_Meses`, hoja `FORMATO OBL`, parámetros `C7`, `D6`, `D7`, resultado `G6`; aportantes por query Teradata Formato 491; TRM | USD por mil aportantes |
| 62 | `gastos / (aportesRecibidos136 / TRM) * 100` | Gastos de `CUENTAS`; aportes de Formato 136; TRM | Porcentaje |
| 63 | `(patrimonioBaseMesMMCop / TRM) / fila28 * 100` | `Plantilla AIOS-probable`, base mes; TRM; fila 28 | Porcentaje |
| 64 | `patrimonioUsd / afiliados * 1,000,000` | Fila 50 y afiliados totales por query Teradata Formato 491 | USD por afiliado |
| 65 | `resultadoNeto / comisiones * 100` | `CUENTAS`, filas 54 y 51 | Porcentaje |
| 66 | `resultadoNeto / patrimonioUsd * 100` | `CUENTAS` y fila 50 | Porcentaje |
| 67 | `gastos / afiliados * 1,000,000` | `CUENTAS` y afiliados totales por query Teradata Formato 491 | Valor por afiliado |
| 68 | `comisiones / aportantes * 1,000,000` | `CUENTAS` y aportantes por query Teradata Formato 491 | Valor por aportante |
| 69 | `admon / fila61` | `CUENTAS` y fila 61 | Ratio |
| 70 | `16` | Constante | Valor crudo |
| 71 | `promedio(col_obl, por_obl, pro_obl, ska_obl) * 100` | `TrimestralData.comisionesPct` | Porcentaje |
| 72 | `0` | Constante | Valor crudo |
| 73 | `0` | Constante | Valor crudo |
| 74 | `(3 - fila71) * 0.25` | Fila 71 y constantes | Porcentaje / puntos |
| 75 | `(3 - fila71) * 0.75` | Fila 71 y constantes | Porcentaje / puntos |
| 76 | `0` | Constante | Valor crudo |
| 77 | `comisiones` | `Plantilla AIOS-probable`, hoja `CUENTAS`, `E13` | COP / unidad de plantilla |
| 78 | `fila28` | Reutiliza fondos administrados de fila 28 | MM USD |
| 79 | `fila77 / fila78` | Filas 77 y 78 | Ratio |
| 80 | `año(fechaCorte) - 1994` | Fecha de corte | Años |
| 81 | Sin información disponible | No se encontró mapeo o fuente implementada para esta fila en el generador semestral | No aplica |
| 82 | `rentabilidad nominal 10 años` | `RentabilidadService`, NAV de `Valores_Fondo_Moder` e IPC de `Rent_Vr_Uni_Moderado` | Porcentaje |
| 83 | `rentabilidad real 10 años` | `RentabilidadService`, nominal 10 años e IPC | Porcentaje |
| 84 | `rentabilidad nominal 5 años` | `RentabilidadService`, NAV de `Valores_Fondo_Moder` | Porcentaje |
| 85 | `rentabilidad real 5 años` | `RentabilidadService`, IPC de `Rent_Vr_Uni_Moderado` | Porcentaje |
| 86 | `rentabilidad nominal 3 años` | `RentabilidadService`, NAV de `Valores_Fondo_Moder` | Porcentaje |
| 87 | `rentabilidad real 3 años` | `RentabilidadService`, IPC de `Rent_Vr_Uni_Moderado` | Porcentaje |
| 88 | `rentabilidad nominal 1 año` | `RentabilidadService`, NAV de `Valores_Fondo_Moder` | Porcentaje |
| 89 | `rentabilidad real 1 año` | `RentabilidadService`, IPC de `Rent_Vr_Uni_Moderado` | Porcentaje |


## 5. Diccionario conceptual de datos semestrales

Esta sección complementa la tabla técnica anterior. La tabla técnica indica **de dónde sale el número**; este diccionario indica **qué significa el indicador** y cómo debe interpretarse funcionalmente.

| Fila | Nombre conceptual | ¿Qué representa? | Interpretación |
|---:|---|---|---|
| 3 | Afiliados totales | Total de personas afiliadas al sistema de fondos de pensiones. | Tamaño de la población cubierta por el sistema. |
| 4 | Afiliados menores de 30 años (%) | Participación de afiliados jóvenes dentro del total de afiliados. | Mide qué proporción del sistema está en edades tempranas de acumulación. |
| 5 | Afiliados de 30 a 44 años (%) | Participación de afiliados en edad laboral media. | Permite ver la concentración de afiliados en etapa de acumulación consolidada. |
| 6 | Afiliados de 45 a 59 años (%) | Participación de afiliados cercanos a edades previas al retiro. | Indica presión futura de maduración del sistema. |
| 7 | Afiliados mayores de 60 años (%) | Participación de afiliados de mayor edad. | Mide la proporción de población próxima o posterior a edades típicas de retiro. |
| 8 | Total de distribución por edad | Base porcentual de referencia. | Sirve como control conceptual de que los grupos de edad componen el 100%. |
| 9 | Afiliados en miles | Total de afiliados expresado en miles. | Facilita comparaciones internacionales o gráficas con magnitudes más manejables. |
| 10 | Mujeres afiliadas (%) | Participación de mujeres dentro del total de afiliados. | Mide composición por género del sistema. |
| 11 | Aportantes | Personas que realizaron aportes al sistema. | Aproxima la población activa que contribuye efectivamente. |
| 12 | Afiliados / PEA (%) | Afiliados como proporción de la población económicamente activa. | Mide cobertura del sistema frente al mercado laboral potencial. |
| 13 | Aportantes / PEA (%) | Aportantes como proporción de la población económicamente activa. | Mide cobertura contributiva efectiva frente al mercado laboral. |
| 14 | Aportantes / afiliados (%) | Proporción de afiliados que aportan. | Indica densidad contributiva o actividad efectiva de afiliados. |
| 15 | Salario mínimo en USD | Salario mínimo colombiano convertido a dólares. | Permite comparación internacional del ingreso mínimo de referencia. |
| 16 | Pensionados totales | Número total de pensionados reportados. | Tamaño de la población pensionada atendida por el sistema. |
| 17 | Pensionados por invalidez (%) | Proporción de pensionados cuya modalidad es invalidez. | Mide composición de beneficios por riesgo de invalidez. |
| 18 | Pensionados por vejez (%) | Proporción de pensionados cuya modalidad es vejez. | Mide peso de las pensiones asociadas a retiro por edad. |
| 19 | Pensionados por sobrevivencia (%) | Proporción de pensionados por sobrevivencia. | Mide peso de beneficios derivados para beneficiarios. |
| 25 | Afiliados fallecidos en miles | Total de afiliados fallecidos del sistema calculado por Formato 493, hoja `Fallecidos`, celda `M11`, con `B11` igual a la fecha de corte y `D4=99`. | Dimensiona en miles el flujo de afiliados fallecidos acumulado en el periodo definido por la fila 11 del insumo. |
| 26 | Traspasos del sistema | Total de movimientos de traslado entre administradoras/fondos. | Indica movilidad de afiliados dentro del sistema. |
| 27 | Traspasos / afiliados (%) | Traspasos respecto al total de afiliados. | Mide intensidad relativa de movilidad en el sistema. |
| 28 | Fondos administrados | Valor total del portafolio administrado por los fondos de pensiones, convertido a millones de USD. | Indica tamaño financiero del sistema pensional. |
| 29 | Fondos administrados / PIB (%) | Proporción del valor total de los fondos frente al Producto Interno Bruto del país. | Responde qué tan grande es el sistema de fondos de pensiones frente al tamaño total de la economía. |
| 30 | Portafolio total en USD | Valor total del portafolio de referencia convertido por TRM. | Base monetaria para analizar composición del portafolio. |
| 31 | Deuda gubernamental local (%) | Proporción del portafolio total invertida en deuda gubernamental local o interna. | Mide exposición del portafolio a deuda pública interna. |
| 32 | Depósitos / efectivo locales (%) | Proporción local en instrumentos de liquidez o efectivo, según clasificación de límites. | Mide liquidez local dentro del portafolio. |
| 33 | Deuda no financiera local (%) | Proporción invertida en emisores no financieros locales. | Indica exposición a deuda corporativa o instrumentos locales no financieros. |
| 34 | Acciones locales (%) | Proporción invertida en renta variable local. | Mide exposición del portafolio al mercado accionario colombiano. |
| 35 | Fondos locales (%) | Proporción invertida en fondos o vehículos locales. | Mide uso de vehículos colectivos locales dentro del portafolio. |
| 36 | Categoría local no usada | Valor constante cero. | Reserva de plantilla sin dato activo. |
| 37 | Deuda gubernamental exterior (%) | Proporción del portafolio invertida en deuda gubernamental extranjera. | Mide exposición soberana internacional. |
| 38 | Depósitos / efectivo exterior (%) | Proporción en liquidez o efectivo en el exterior. | Mide liquidez internacional del portafolio. |
| 39 | Deuda no financiera exterior (%) | Proporción en deuda privada o no financiera del exterior. | Mide exposición crediticia internacional no soberana. |
| 40 | Acciones exterior (%) | Proporción en renta variable extranjera. | Mide exposición accionaria internacional. |
| 41 | Fondos exterior (%) | Proporción invertida en fondos o vehículos del exterior. | Mide diversificación internacional mediante vehículos colectivos. |
| 42 | Referencia normativa | Valor fijo usado por la plantilla. | Dato de control o referencia no derivado de insumos. |
| 43 | Otros activos (%) | Proporción del portafolio en otras categorías. | Completa la clasificación de activos no cubierta por rubros anteriores. |
| 44 | Suma de rubros exteriores seleccionados (%) | Agregado de celdas seleccionadas de límites exteriores. | Resume exposición exterior de categorías específicas. |
| 45 | Fondos / deuda gubernamental total (%) | Fondos administrados frente a deuda gubernamental total en USD. | Mide el tamaño relativo del sistema pensional respecto al saldo de deuda pública. |
| 46 | Referencia fija | Constante de plantilla. | Dato de control o referencia. |
| 47 | Participación Protección + Porvenir (%) | Participación conjunta de Protección y Porvenir sobre el total del sistema. | Mide concentración de mercado en las dos administradoras indicadas. |
| 48 | Activos en USD | Activos contables convertidos a dólares. | Mide tamaño del balance por el lado de activos. |
| 49 | Pasivos en USD | Pasivos contables convertidos a dólares. | Mide obligaciones del balance en dólares. |
| 50 | Patrimonio en USD | Activos menos pasivos, convertido a dólares. | Mide valor patrimonial contable. |
| 51 | Comisiones | Ingresos por comisiones según plantilla contable. | Mide ingresos operacionales asociados a administración. |
| 52 | Gastos | Gastos reportados en la plantilla contable. | Mide egresos operativos/administrativos. |
| 53 | Resultado operacional | Resultado de operación antes de resultado neto. | Mide desempeño operativo. |
| 54 | Resultado neto | Resultado final neto. | Mide utilidad o pérdida final del periodo. |
| 55 | Gastos de administración | Rubro administrativo seleccionado. | Mide gasto administrativo relevante. |
| 56 | Comisión 511500 en USD | Cuenta 511500 convertida a dólares. | Mide el rubro específico de comisiones en moneda comparable. |
| 57 | Comisión 511527 en USD | Cuenta 511527 convertida a dólares. | Mide afiliaciones a fondos de pensiones en moneda comparable. |
| 58 | Comisiones 511500 + 511527 en USD | Suma de las dos cuentas anteriores convertida a dólares. | Resume rubros de comisión/afiliación solicitados. |
| 59 | Otros gastos operacionales en USD | Suma de cuentas de beneficios, honorarios, cambios, impuestos, arrendamientos, contribuciones, seguros, mantenimiento, adecuación, deterioro y multas, convertida a dólares. | Mide gastos operacionales seleccionados distintos de los rubros principales. |
| 60 | Gasto de operación 510000 en USD | Cuenta 510000 convertida a dólares. | Mide gasto operacional agregado en moneda comparable. |
| 61 | Aportes recibidos por aportante | Aportes recibidos convertidos a USD y normalizados por aportantes en miles. | Mide intensidad de aportes recibidos por población aportante. |
| 62 | Gastos / aportes recibidos (%) | Gastos frente a aportes recibidos convertidos a USD. | Mide carga de gastos sobre los aportes recibidos. |
| 63 | Patrimonio / fondos administrados (%) | Patrimonio de base mes en USD frente a fondos administrados. | Mide respaldo patrimonial relativo al tamaño de fondos. |
| 64 | Patrimonio por afiliado | Patrimonio en USD dividido por afiliados y escalado. | Mide patrimonio relativo por afiliado. |
| 65 | Resultado neto / comisiones (%) | Resultado neto frente a comisiones. | Mide rentabilidad o margen sobre ingresos por comisiones. |
| 66 | Resultado neto / patrimonio (%) | Resultado neto frente al patrimonio. | Aproxima retorno sobre patrimonio. |
| 67 | Gastos por afiliado | Gastos divididos por afiliados y escalados. | Mide costo promedio asociado a afiliados. |
| 68 | Comisiones por aportante | Comisiones divididas por aportantes y escaladas. | Mide ingreso promedio por aportante. |
| 69 | Administración / aportes por aportante | Gasto administrativo frente al indicador de fila 61. | Mide carga administrativa respecto al flujo de aportes normalizado. |
| 70 | Referencia fija | Constante de plantilla. | Dato de control o referencia. |
| 71 | Comisión promedio obligatoria (%) | Promedio de comisiones obligatorias de administradoras seleccionadas. | Mide costo promedio de comisión obligatoria. |
| 72 | Aporte adicional trabajador | Constante cero. | Rubro de plantilla sin dato activo. |
| 73 | Aporte adicional empleador | Constante cero. | Rubro de plantilla sin dato activo. |
| 74 | Aporte trabajador | Parte trabajador de la diferencia entre 3 y comisión promedio. | Estima distribución del aporte residual hacia trabajador. |
| 75 | Aporte empleador | Parte empleador de la diferencia entre 3 y comisión promedio. | Estima distribución del aporte residual hacia empleador. |
| 76 | Referencia fija | Constante cero. | Rubro de plantilla sin dato activo. |
| 77 | Comisiones | Reutiliza comisiones contables. | Base para medir comisiones respecto a fondos. |
| 78 | Fondos administrados | Reutiliza fila 28. | Base financiera del sistema. |
| 79 | Comisiones / fondos | Comisiones frente a fondos administrados. | Mide peso de comisiones sobre activos administrados. |
| 80 | Años desde 1994 | Diferencia entre año de corte y 1994. | Mide antigüedad del régimen o periodo de referencia. |
| 82 | Rentabilidad nominal 10 años | Retorno nominal anualizado de largo plazo. | Mide desempeño sin descontar inflación. |
| 83 | Rentabilidad real 10 años | Retorno real anualizado de largo plazo. | Mide desempeño descontando inflación. |
| 84 | Rentabilidad nominal 5 años | Retorno nominal anualizado de mediano plazo. | Mide desempeño a cinco años sin inflación. |
| 85 | Rentabilidad real 5 años | Retorno real anualizado de mediano plazo. | Mide desempeño a cinco años descontando inflación. |
| 86 | Rentabilidad nominal 3 años | Retorno nominal anualizado de mediano/corto plazo. | Mide desempeño a tres años sin inflación. |
| 87 | Rentabilidad real 3 años | Retorno real anualizado de mediano/corto plazo. | Mide desempeño a tres años descontando inflación. |
| 88 | Rentabilidad nominal 1 año | Retorno nominal anual de corto plazo. | Mide desempeño reciente sin inflación. |
| 89 | Rentabilidad real 1 año | Retorno real anual de corto plazo. | Mide desempeño reciente descontando inflación. |

### 5.1 Fórmulas conceptuales clave

$$
\text{Afiliados / PEA} = \frac{\text{Afiliados}}{\text{Población Económicamente Activa}} \times 100
$$

$$
\text{Aportantes / PEA} = \frac{\text{Aportantes}}{\text{Población Económicamente Activa}} \times 100
$$

$$
\text{Traspasos / Afiliados} = \frac{\text{Traspasos del sistema}}{\text{Afiliados}} \times 100
$$

$$
\text{Fondos / Deuda pública total} = \frac{\text{Fondos administrados}}{\text{Deuda gubernamental total}} \times 100
$$

$$
\text{Participación Protección + Porvenir} = \frac{\text{Fondos de Protección} + \text{Fondos de Porvenir}}{\text{Fondos totales del sistema}} \times 100
$$

$$
\text{Resultado neto / Patrimonio} = \frac{\text{Resultado neto}}{\text{Patrimonio}} \times 100
$$

$$
\text{Comisiones / Fondos} = \frac{\text{Comisiones}}{\text{Fondos administrados}} \times 100
$$

## 6. Explicación conceptual detallada por archivo

Las siguientes subsecciones usan un patrón completo para cada dato: **qué representa**, **interpretación**, **fórmula conceptual** y, cuando aplica, **fórmula implementada**. La fuente técnica exacta se conserva en las tablas anteriores.

### 6.1 Archivo mensual (`Boletin_AIOS MENSUAL.xlsx`)

#### Columna B: Afiliados totales

**¿Qué representa?**

La columna B del archivo mensual corresponde al número total de personas afiliadas al sistema de fondos de pensiones.

**Interpretación**

Permite dimensionar la cobertura total del sistema y sirve como denominador para indicadores de composición.

**Fórmula conceptual**

$$
\text{Afiliados} = \text{Hombres afiliados} + \text{Mujeres afiliadas}
$$

#### Columna C: Aportantes

**¿Qué representa?**

La columna C del archivo mensual corresponde al número de afiliados que realizaron aportes en el periodo reportado.

**Interpretación**

Mide la base activa que efectivamente contribuye al sistema.

**Fórmula conceptual**

$$
\text{Aportantes} = \text{Afiliados con aporte registrado}
$$

#### Columna D: Traspasos del sistema

**¿Qué representa?**

La columna D del archivo mensual corresponde al total de movimientos de traslado entre administradoras o fondos.

**Interpretación**

Mide movilidad o rotación de afiliados dentro del sistema.

**Fórmula conceptual**

$$
\text{Traspasos} = \text{Total de traslados reportados}
$$

#### Columna E: Valor del fondo en USD

**¿Qué representa?**

La columna E del archivo mensual corresponde al valor del fondo administrado convertido a dólares.

**Interpretación**

Permite comparar el tamaño financiero del portafolio en una moneda común.

**Fórmula conceptual**

$$
\text{Valor del fondo (USD)} = \frac{\text{Valor del fondo en COP}}{\text{TRM}}
$$

#### Columna F: Total de límites en USD

**¿Qué representa?**

La columna F del archivo mensual corresponde al total de referencia de límites convertido a dólares.

**Interpretación**

Sirve como base monetaria para revisar composición de límites de inversión.

**Fórmula conceptual**

$$
\text{Total límites (USD)} = \frac{\text{Total límites en COP}}{\text{TRM}}
$$

#### Columna G: Deuda gubernamental local (%)

**¿Qué representa?**

La columna G del archivo mensual corresponde a la proporción del portafolio invertida en deuda pública interna.

**Interpretación**

Mide exposición a deuda soberana local.

**Fórmula conceptual**

$$
\text{Deuda gubernamental local} = \frac{\text{Deuda gubernamental local}}{\text{Portafolio total}} \times 100
$$

#### Columna H: Depósitos y efectivo local (%)

**¿Qué representa?**

La columna H del archivo mensual corresponde a la proporción del portafolio en liquidez local.

**Interpretación**

Mide el peso de instrumentos líquidos locales dentro del portafolio.

**Fórmula conceptual**

$$
\text{Efectivo local} = \frac{\text{Depósitos y efectivo locales}}{\text{Portafolio total}} \times 100
$$

#### Columna I: Deuda no financiera local (%)

**¿Qué representa?**

La columna I del archivo mensual corresponde a la proporción invertida en deuda local de emisores no financieros.

**Interpretación**

Mide exposición a crédito corporativo local.

**Fórmula conceptual**

$$
\text{Deuda no financiera local} = \frac{\text{Deuda no financiera local}}{\text{Portafolio total}} \times 100
$$

#### Columna J: Acciones locales (%)

**¿Qué representa?**

La columna J del archivo mensual corresponde a la proporción del portafolio invertida en renta variable local.

**Interpretación**

Mide exposición al mercado accionario colombiano.

**Fórmula conceptual**

$$
\text{Acciones locales} = \frac{\text{Acciones locales}}{\text{Portafolio total}} \times 100
$$

#### Columna K: Fondos locales (%)

**¿Qué representa?**

La columna K del archivo mensual corresponde a la proporción invertida en vehículos colectivos locales.

**Interpretación**

Mide uso de fondos o vehículos locales dentro del portafolio.

**Fórmula conceptual**

$$
\text{Fondos locales} = \frac{\text{Fondos locales}}{\text{Portafolio total}} \times 100
$$

#### Columna L: Inversiones del exterior seleccionadas (%)

**¿Qué representa?**

La columna L del archivo mensual corresponde al agregado de categorías de inversión exterior seleccionadas.

**Interpretación**

Resume diversificación internacional en rubros específicos.

**Fórmula conceptual**

$$
\text{Exterior seleccionado} = \frac{\text{Rubros exteriores seleccionados}}{\text{Portafolio total}} \times 100
$$

#### Columna M: Otros activos (%)

**¿Qué representa?**

La columna M del archivo mensual corresponde a categorías de portafolio no clasificadas en rubros principales.

**Interpretación**

Completa la composición del portafolio y ayuda a identificar saldos residuales.

**Fórmula conceptual**

$$
\text{Otros activos} = \frac{\text{Otros activos}}{\text{Portafolio total}} \times 100
$$

#### Columna N: Rentabilidad nominal 1 año (%)

**¿Qué representa?**

La columna N del archivo mensual corresponde al retorno anual sin descontar inflación.

**Interpretación**

Mide desempeño financiero observado en términos nominales.

**Fórmula conceptual**

$$
\text{Rentabilidad nominal} = \left(\frac{\text{Valor final}}{\text{Valor inicial}} - 1\right) \times 100
$$

#### Columna O: Rentabilidad real 1 año (%)

**¿Qué representa?**

La columna O del archivo mensual corresponde al retorno anual descontando inflación.

**Interpretación**

Mide ganancia de poder adquisitivo del portafolio.

**Fórmula conceptual**

$$
\text{Rentabilidad real} = \left(\frac{1+\text{Rentabilidad nominal}}{1+\text{Inflación}} - 1\right) \times 100
$$

#### Columna P: Referencia fija

**¿Qué representa?**

La columna P del archivo mensual corresponde a un valor constante de la plantilla.

**Interpretación**

Funciona como control o referencia operacional sin fuente externa.

**Fórmula conceptual**

$$
\text{Referencia} = 4
$$

#### Columna Q: Concentración de afiliados (personas)

**¿Qué representa?**

La columna Q del archivo mensual corresponde a la concentración de afiliados, es decir, personas afiliadas. Se calcula con query sobre `PROD_DWH_CONSULTA.FORMATO491`, no con el Excel 491.

**Interpretación**

Mide peso relativo de fondos/administradoras seleccionadas dentro del sistema.

**Fórmula conceptual**

$$
\text{Concentración afiliados} = \frac{\text{Afiliados de las dos AFP con más afiliados}}{\text{Afiliados totales del sistema}} \times 100
$$

#### Columna R: Participación Protección + Porvenir (%)

**¿Qué representa?**

La columna R del archivo mensual corresponde a la participación conjunta de Protección y Porvenir sobre el total del sistema.

**Interpretación**

Mide concentración de mercado en esas administradoras.

**Fórmula conceptual**

$$
\text{Participación} = \frac{\text{Protección} + \text{Porvenir}}{\text{Sistema total}} \times 100
$$

#### Columna S: TRM

**¿Qué representa?**

La columna S del archivo mensual corresponde a la tasa de cambio usada para convertir COP a USD.

**Interpretación**

Permite expresar valores monetarios en una moneda comparable.

**Fórmula conceptual**

$$
\text{USD} = \frac{\text{COP}}{\text{TRM}}
$$

### 6.2 Archivo trimestral (`Boletin_AIOS TRIMESTRAL.xlsx`)

#### Hoja `afiliados`: Afiliados por administradora y fondo

**¿Qué representa?**

La hoja `afiliados` del archivo trimestral corresponde al número de afiliados distribuido por administradora y tipo de fondo.

**Interpretación**

Permite analizar composición, participación y concentración de afiliados entre entidades y fondos.

**Fórmula conceptual**

$$
\text{Afiliados por grupo} = \sum \text{Afiliados del grupo reportado}
$$

#### Hoja `aportantes`: Aportantes por administradora

**¿Qué representa?**

La hoja `aportantes` del archivo trimestral corresponde al número de personas que aportan por administradora.

**Interpretación**

Mide actividad contributiva efectiva y permite comparar entidades.

**Fórmula conceptual**

$$
\text{Aportantes por administradora} = \sum \text{Aportantes reportados}
$$

#### Hoja `colombia`: Saldos de Colombia por fondo

**¿Qué representa?**

La hoja `colombia` del archivo trimestral corresponde a saldos o fondos administrados asociados a Colombia, por administradora y fondo.

**Interpretación**

Mide el tamaño de los recursos administrados por segmento.

**Fórmula conceptual**

$$
\text{Saldo Colombia} = \text{Valor reportado por fondo/administradora}
$$

#### Hoja `traspasos`: Traspasos trimestrales

**¿Qué representa?**

La hoja `traspasos` del archivo trimestral corresponde a movimientos de traslado por administradora durante el periodo.

**Interpretación**

Mide movilidad relativa de afiliados y competencia por traslados.

**Fórmula conceptual**

$$
\text{Traspasos trimestrales} = \sum \text{Traslados del trimestre}
$$

#### Hoja `gastos`: Gastos en USD

**¿Qué representa?**

La hoja `gastos` del archivo trimestral corresponde a gastos netos convertidos a dólares.

**Interpretación**

Permite comparar egresos entre administradoras usando una moneda homogénea.

**Fórmula conceptual**

$$
\text{Gastos (USD)} = \frac{\text{Gastos netos en COP}}{\text{TRM}}
$$

#### Hoja `promotores`: Promotores

**¿Qué representa?**

La hoja `promotores` del archivo trimestral corresponde al número de promotores o fuerza comercial cuando exista fuente disponible.

**Interpretación**

Mide capacidad comercial o red de atención asociada a las administradoras.

**Fórmula conceptual**

$$
\text{Promotores} = \text{Cantidad reportada de promotores}
$$

#### Hoja `rentabilidad`: Rentabilidad nominal y real

**¿Qué representa?**

La hoja `rentabilidad` del archivo trimestral corresponde al rendimiento de los fondos por administradora/fondo.

**Interpretación**

Mide desempeño financiero con y sin efecto de inflación.

**Fórmula conceptual**

$$
\text{Rentabilidad real} = \left(\frac{1+\text{Rentabilidad nominal}}{1+\text{Inflación}} - 1\right) \times 100
$$

#### Hoja `comisiones`: Comisiones obligatorias

**¿Qué representa?**

La hoja `comisiones` del archivo trimestral corresponde al porcentaje de comisión cobrado por las administradoras para aportes obligatorios.

**Interpretación**

Mide costo para el afiliado/aportante por administración de recursos.

**Fórmula conceptual**

$$
\text{Comisión promedio} = \frac{\sum \text{Comisiones por administradora}}{\text{Número de administradoras}}
$$

### 6.3 Archivo semestral (`Semestral_Colombia.xlsx` / `semestral.xlsx`)

#### Fila 3: Afiliados totales

**¿Qué representa la fila 3?**

La fila 3 del archivo `Semestral_Colombia.xlsx` corresponde al número total de personas afiliadas al sistema de fondos de pensiones.

**Interpretación económica u operativa**

Este dato dimensiona la cobertura total del sistema y sirve como base para calcular composiciones por edad, género y actividad contributiva.

**Fórmula conceptual**

$$
\text{Afiliados totales} = \text{Hombres afiliados} + \text{Mujeres afiliadas}
$$

**Fórmula implementada**

$$
\text{Fila 3} = \text{mensual.afiliados}
$$

Se obtiene desde query Teradata sobre `PROD_DWH_CONSULTA.FORMATO491`, sumando `TOTAL_AFILIADOS_TOTAL` para `RENGLON=999` y fondos 1000/5000/6000/7000/8000.

#### Fila 4: Afiliados menores de 30 años (%)

**¿Qué representa la fila 4?**

La fila 4 corresponde a la proporción de afiliados menores de 30 años dentro del total de afiliados.

**Interpretación económica u operativa**

Indica qué tan joven es la base de afiliados y qué proporción se encuentra en etapas tempranas de acumulación pensional.

**Fórmula conceptual**

$$
\text{Menores de 30} = \frac{\text{Afiliados menores de 30}}{\text{Afiliados totales}} \times 100
$$

**Fórmula implementada**

$$
\text{Fila 4} = \frac{\text{mensual.afiliadosMenor30}}{\text{mensual.afiliados}} \times 100
$$

Los afiliados menores de 30 provienen de Formato 491, `informe de prensa`, celdas `C81 + D81`.

#### Fila 5: Afiliados de 30 a 44 años (%)

**¿Qué representa la fila 5?**

La fila 5 corresponde a la proporción de afiliados entre 30 y 44 años dentro del total de afiliados.

**Interpretación económica u operativa**

Permite observar la concentración de afiliados en una etapa laboral media, usualmente asociada a acumulación pensional estable.

**Fórmula conceptual**

$$
\text{Afiliados 30 a 44} = \frac{\text{Afiliados de 30 a 44}}{\text{Afiliados totales}} \times 100
$$

**Fórmula implementada**

$$
\text{Fila 5} = \frac{\text{mensual.afiliados30a44}}{\text{mensual.afiliados}} \times 100
$$

El numerador proviene de Formato 491, `informe de prensa`, celdas `C82 + D82`.

#### Fila 6: Afiliados de 45 a 59 años (%)

**¿Qué representa la fila 6?**

La fila 6 corresponde a la proporción de afiliados entre 45 y 59 años dentro del total.

**Interpretación económica u operativa**

Ayuda a anticipar la maduración del sistema y el peso de afiliados próximos a edades de retiro.

**Fórmula conceptual**

$$
\text{Afiliados 45 a 59} = \frac{\text{Afiliados de 45 a 59}}{\text{Afiliados totales}} \times 100
$$

**Fórmula implementada**

$$
\text{Fila 6} = \frac{\text{mensual.afiliados45a59}}{\text{mensual.afiliados}} \times 100
$$

El numerador proviene de Formato 491, `informe de prensa`, celdas `C83 + D83`.

#### Fila 7: Afiliados mayores de 60 años (%)

**¿Qué representa la fila 7?**

La fila 7 corresponde a la proporción de afiliados mayores de 60 años dentro del total.

**Interpretación económica u operativa**

Mide la presencia de población afiliada próxima o posterior a edades típicas de retiro.

**Fórmula conceptual**

$$
\text{Afiliados mayores de 60} = \frac{\text{Afiliados mayores de 60}}{\text{Afiliados totales}} \times 100
$$

**Fórmula implementada**

$$
\text{Fila 7} = \frac{\text{mensual.afiliadosMayor60}}{\text{mensual.afiliados}} \times 100
$$

El numerador proviene de Formato 491, `informe de prensa`, celdas `C84 + D84`.

#### Fila 8: Total de distribución por edad

**¿Qué representa la fila 8?**

La fila 8 representa el total porcentual de los grupos de edad.

**Interpretación económica u operativa**

Funciona como control conceptual: la suma de las participaciones por edad debe explicar el total de afiliados.

**Fórmula conceptual**

$$
\text{Total grupos de edad} = 100
$$

**Fórmula implementada**

$$
\text{Fila 8} = 100
$$

Es una constante de la plantilla.

#### Fila 9: Afiliados en miles

**¿Qué representa la fila 9?**

La fila 9 muestra el total de afiliados expresado en miles de personas.

**Interpretación económica u operativa**

Reduce la escala del dato poblacional y facilita comparación con otros países o series históricas.

**Fórmula conceptual**

$$
\text{Afiliados en miles} = \frac{\text{Afiliados totales}}{1000}
$$

**Fórmula implementada**

$$
\text{Fila 9} = \frac{\text{mensual.afiliados}}{1000}
$$

Usa el total de afiliados de la fila 3.

#### Fila 10: Mujeres afiliadas (%)

**¿Qué representa la fila 10?**

La fila 10 corresponde a la participación de mujeres dentro del total de afiliados.

**Interpretación económica u operativa**

Permite analizar composición por género del sistema pensional.

**Fórmula conceptual**

$$
\text{Mujeres afiliadas} = \frac{\text{Mujeres afiliadas}}{\text{Afiliados totales}} \times 100
$$

**Fórmula implementada**

$$
\text{Fila 10} = \frac{\text{mensual.mujeres}}{\text{mensual.afiliados}} \times 100
$$

Mujeres se lee por query Teradata sobre `PROD_DWH_CONSULTA.FORMATO491`, sumando `TOTAL_AFILIADOS_M` para `RENGLON = '999'` y los fondos obligatorios `1000`, `5000`, `6000`, `7000` y `8000`. El denominador `mensual.afiliados` usa el total de afiliados por query (`SUM(TOTAL_AFILIADOS_TOTAL)`) con los mismos filtros.

#### Fila 11: Aportantes

**¿Qué representa la fila 11?**

La fila 11 corresponde al número de personas que realizaron aportes al sistema.

**Interpretación económica u operativa**

Representa la base activa que efectivamente contribuye y no solo la población afiliada.

**Fórmula conceptual**

$$
\text{Aportantes} = \text{Personas con aporte registrado}
$$

**Fórmula implementada**

$$
\text{Fila 11} = \text{mensual.aportantes}
$$

Se lee por query Teradata desde `PROD_DWH_CONSULTA.FORMATO491`, sumando `TOTAL_AFILIADOS_COTIZANTES` para `RENGLON=999`, fondos 1000/5000/6000/7000/8000 y sin filtro por `CODIGO_ENTIDAD` en el semestral.

#### Fila 12: Afiliados / PEA (%)

**¿Qué representa la fila 12?**

La fila 12 muestra los afiliados como proporción de la Población Económicamente Activa.

**Interpretación económica u operativa**

Mide la cobertura del sistema frente al mercado laboral potencial.

**Fórmula conceptual**

$$
\text{Afiliados sobre PEA} = \frac{\text{Afiliados}}{\text{PEA}} \times 100
$$

**Fórmula implementada**

$$
\text{Fila 12} = \frac{\text{mensual.afiliados}}{\text{mensual.pea}} \times 100
$$

La PEA proviene de `PIB_PEA_TRM_DG`.

#### Fila 13: Aportantes / PEA (%)

**¿Qué representa la fila 13?**

La fila 13 muestra los aportantes como proporción de la Población Económicamente Activa.

**Interpretación económica u operativa**

Mide la cobertura contributiva efectiva del sistema frente al mercado laboral.

**Fórmula conceptual**

$$
\text{Aportantes sobre PEA} = \frac{\text{Aportantes}}{\text{PEA}} \times 100
$$

**Fórmula implementada**

$$
\text{Fila 13} = \frac{\text{mensual.aportantes}}{\text{mensual.pea}} \times 100
$$

Combina aportantes por query Teradata del Formato 491 con PEA de `PIB_PEA_TRM_DG`.

#### Fila 14: Aportantes / afiliados (%)

**¿Qué representa la fila 14?**

La fila 14 corresponde al porcentaje de afiliados que aportan efectivamente.

**Interpretación económica u operativa**

Es un indicador de densidad contributiva o actividad efectiva de la base afiliada.

**Fórmula conceptual**

$$
\text{Aportantes sobre afiliados} = \frac{\text{Aportantes}}{\text{Afiliados}} \times 100
$$

**Fórmula implementada**

$$
\text{Fila 14} = \frac{\text{mensual.aportantes}}{\text{mensual.afiliados}} \times 100
$$

Los aportantes y afiliados provienen de queries Teradata del Formato 491.

#### Fila 15: Salario mínimo en USD

**¿Qué representa la fila 15?**

La fila 15 corresponde al salario mínimo colombiano convertido a dólares.

**Interpretación económica u operativa**

Permite comparar internacionalmente un ingreso laboral de referencia.

**Fórmula conceptual**

$$
\text{Salario mínimo en USD} = \frac{\text{Salario mínimo en COP}}{\text{TRM}}
$$

**Fórmula implementada**

$$
\text{Fila 15} = \text{mensual.smColombiaUsd}
$$

El salario mínimo ponderado en COP ya no se toma de la celda `E8` del Excel del Formato 491. La aplicación lee el salario mínimo oficial del año desde `SalarioMinimo.csv` y ejecuta una consulta a `PROD_DWH_CONSULTA.FORMATO491` que pondera los afiliados por rangos de IBC (`1`, `2`, `3`, `4`, `8`, `12`, `16`, `20` y `25` salarios mínimos) con las reglas de `UNIDAD_CAPTURA` y fondo indicadas para moderado, conservador y mayor riesgo. La TRM continúa saliendo de `PIB_PEA_TRM_DG`.

#### Fila 16: Pensionados totales

**¿Qué representa la fila 16?**

La fila 16 corresponde al total de pensionados reportados.

**Interpretación económica u operativa**

Dimensiona la población beneficiaria que recibe pensión dentro del sistema.

**Fórmula conceptual**

$$
\text{Pensionados totales} = \text{Total de pensionados reportados}
$$

**Fórmula implementada**

$$
\text{Fila 16} = \text{total pensionados de Formato 495}
$$

Se lee de Formato 495, hoja `TOTAL PENSIONADOS`, con parámetro `B4` y valor en columna `I`.

#### Fila 17: Pensionados por invalidez (%)

**¿Qué representa la fila 17?**

La fila 17 corresponde a la proporción de pensionados por invalidez.

**Interpretación económica u operativa**

Mide el peso relativo de las pensiones originadas por invalidez.

**Fórmula conceptual**

$$
\text{Invalidez} = \frac{\text{Pensionados por invalidez}}{\text{Pensionados totales}} \times 100
$$

**Fórmula implementada**

$$
\text{Fila 17} = \frac{\text{por Entidad!BI62}}{\text{Fila 16}}
$$

Se obtiene de Formato 495, hoja `por Entidad`, celda `BI62`.

#### Fila 18: Pensionados por vejez (%)

**¿Qué representa la fila 18?**

La fila 18 corresponde a la proporción de pensionados por vejez.

**Interpretación económica u operativa**

Mide el peso relativo de pensiones asociadas al retiro por edad.

**Fórmula conceptual**

$$
\text{Vejez} = \frac{\text{Pensionados por vejez}}{\text{Pensionados totales}} \times 100
$$

**Fórmula implementada**

$$
\text{Fila 18} = \frac{\text{por Entidad!BH62}}{\text{Fila 16}}
$$

Se obtiene de Formato 495, hoja `por Entidad`, celda `BH62`.

#### Fila 19: Pensionados por sobrevivencia (%)

**¿Qué representa la fila 19?**

La fila 19 corresponde a la proporción de pensionados por sobrevivencia.

**Interpretación económica u operativa**

Mide el peso de beneficios pagados a beneficiarios por sobrevivencia.

**Fórmula conceptual**

$$
\text{Sobrevivencia} = \frac{\text{Pensionados por sobrevivencia}}{\text{Pensionados totales}} \times 100
$$

**Fórmula implementada**

$$
\text{Fila 19} = \frac{\text{por Entidad!BJ62}}{\text{Fila 16}}
$$

Se obtiene de Formato 495, hoja `por Entidad`, celda `BJ62`.

#### Fila 25: Movimiento de afiliados desde Formato 493 (miles)

**¿Qué representa la fila 25?**

La fila 25 toma el total de afiliados fallecidos calculado en el archivo trimestral de referencia `Serie_Formato_493 MOVIMIENTO AFILIADOS.xlsx`. Para el corte solicitado se parametriza la hoja `Fallecidos` escribiendo la fecha en `B11` y el código de entidad `99` en `D4`; luego se lee la celda `M11`.

**Interpretación económica u operativa**

El valor corresponde al total de fallecidos del sistema que calcula el Formato 493 para el rango de fechas de la fila 11 y los rangos de edad/sexo de la hoja `Fallecidos`. Se reporta en miles para mantener la escala del boletín.

**Fórmula conceptual**

$$
\text{Fila 25} = \frac{\text{Formato 493 hoja Fallecidos!M11, con B11 = fecha de corte y D4 = 99}}{1000}
$$

**Fórmula implementada**

`SemestralExcelGenerator` abre `Serie_Formato_493 MOVIMIENTO AFILIADOS.xlsx`, ubica la hoja `Fallecidos`, escribe `fechaCorte` en `B11`, fija `D4=99` para leer el total del sistema, evalúa las fórmulas del libro y escribe `M11 / 1000` en la fila 25 de la salida.

**Interpretación de la fórmula Excel de `M11`**

La fórmula de Excel suma cuatro bloques `SUMAR.SI.CONJUNTO` sobre `Data!U`, uno por cada categoría indicada en `L2`, `M2`, `N2` y `O2` de la hoja `Fallecidos`. En todos los bloques exige que `Data!H` sea igual a `Q2`, que `Data!J` sea igual a la categoría del bloque y que la fecha `Data!D` esté entre `A11` y `B11` inclusive. Si `D4` es diferente de `99`, también filtra `Data!B` por el código de administradora de `D4`; si `D4` es `99`, omite ese filtro y suma el total del sistema. En términos simples: `M11` es la suma de `Data!U` para el periodo `A11:B11`, el concepto de `Q2` y las cuatro categorías `L2:O2`, con filtro opcional de administradora según `D4`.

#### Fila 26: Traspasos del sistema

**¿Qué representa la fila 26?**

La fila 26 corresponde al total de traspasos del sistema.

**Interpretación económica u operativa**

Mide la movilidad total de afiliados entre administradoras o fondos.

**Fórmula conceptual**

$$
\text{Traspasos} = \text{Total de movimientos reportados}
$$

**Fórmula implementada**

$$
\text{Fila 26} = \text{mensual.traspasosSistema}
$$

Proviene de Formato 493 o del lector mensual equivalente.

#### Fila 27: Traspasos / afiliados (%)

**¿Qué representa la fila 27?**

La fila 27 corresponde a la intensidad de traspasos frente al total de afiliados.

**Interpretación económica u operativa**

Permite comparar movilidad relativa independientemente del tamaño del sistema.

**Fórmula conceptual**

$$
\text{Traspasos sobre afiliados} = \frac{\text{Traspasos}}{\text{Afiliados}} \times 100
$$

**Fórmula implementada**

$$
\text{Fila 27} = \frac{\text{mensual.traspasosSistema}}{\text{mensual.afiliados}}
$$

La celda se formatea como porcentaje en Excel.

#### Fila 28: Fondos administrados

**¿Qué representa la fila 28?**

La fila 28 corresponde al valor total de fondos administrados, convertido a millones de dólares.

**Interpretación económica u operativa**

Es el indicador principal del tamaño financiero del sistema de fondos de pensiones.

**Fórmula conceptual**

$$
\text{Fondos administrados} = \frac{\text{Valor de fondos en COP}}{\text{TRM}}
$$

**Fórmula implementada**

$$
\text{Fila 28} = \frac{\text{SISTEMA TOTAL!restot!J14} \times 1000}{\text{TRM} \times 1{,}000{,}000}
$$

La implementación toma `J14` de `SISTEMA TOTAL`, hoja `restot`.

#### Fila 29: Fondos administrados / PIB (%)

**¿Qué representa la fila 29?**

La fila 29 del archivo `Semestral_Colombia.xlsx` corresponde a la proporción del valor total de los fondos de pensiones respecto al Producto Interno Bruto del país, expresada como porcentaje.

**Interpretación económica u operativa**

Este indicador responde qué tan grande es el sistema de fondos de pensiones frente al tamaño total de la economía. Un valor mayor indica que los activos administrados tienen un peso más alto frente a la producción anual del país.

**Fórmula conceptual**

$$
\text{Fondos sobre PIB} = \frac{\text{Valor total de los fondos de pensiones}}{\text{Producto Interno Bruto}} \times 100
$$

**Fórmula implementada**

$$
\text{Fila 29} = \frac{\text{Fila 28}}{\text{PIB semestral en COP} / \text{TRM}}
$$

En la implementación, la fila 28 ya está expresada en millones de USD y el PIB se convierte con TRM para comparar magnitudes.

#### Fila 30: Portafolio total en USD

**¿Qué representa la fila 30?**

La fila 30 corresponde al valor total de referencia del portafolio convertido a dólares.

**Interpretación económica u operativa**

Sirve como base monetaria para analizar composición y exposición del portafolio.

**Fórmula conceptual**

$$
\text{Portafolio total en USD} = \frac{\text{Portafolio total en COP}}{\text{TRM}}
$$

**Fórmula implementada**

$$
\text{Fila 30} = \frac{\text{mensual.total1}}{\text{TRM}}
$$

El total proviene de los datos de límites o composición leídos por el lector mensual.

#### Fila 31: Deuda gubernamental local (%)

**¿Qué representa la fila 31?**

La fila 31 del archivo `Semestral_Colombia.xlsx` corresponde a la proporción del portafolio total de los fondos de pensiones que está invertida en deuda gubernamental local o interna, expresada como porcentaje del total del portafolio.

**Interpretación económica u operativa**

Este indicador muestra la exposición del portafolio administrado a títulos de deuda pública interna. Permite evaluar qué parte de los recursos pensionales financia deuda gubernamental local y qué tan concentrada está la inversión en este tipo de activo.

**Fórmula conceptual**

$$
\text{Deuda gubernamental local} = \frac{\text{Valor invertido en deuda gubernamental local}}{\text{Valor total del portafolio de fondos}} \times 100
$$

**Fórmula implementada**

$$
\text{Fila 31} = \text{LIMITES!AIOS!C4}
$$

El archivo `LIMITES` ya entrega la proporción calculada para esta categoría; por eso se toma directamente el ratio de `C4`.

#### Fila 32: Depósitos y efectivo local (%)

**¿Qué representa la fila 32?**

La fila 32 corresponde a la proporción del portafolio total invertida en depósitos, efectivo o instrumentos de liquidez locales.

**Interpretación económica u operativa**

Permite identificar la exposición del portafolio a depósitos, efectivo o instrumentos de liquidez locales y evaluar la diversificación de inversiones.

**Fórmula conceptual**

$$
\text{Depósitos y efectivo local} = \frac{\text{Valor invertido en depósitos, efectivo o instrumentos de liquidez locales}}{\text{Valor total del portafolio}} \times 100
$$

**Fórmula implementada**

$$
\text{Fila 32} = \text{LIMITES!AIOS!E4}
$$

El valor proviene del archivo `LIMITES`, hoja `AIOS`, o del campo mensual equivalente para esta categoría.

#### Fila 33: Deuda no financiera local (%)

**¿Qué representa la fila 33?**

La fila 33 corresponde a la proporción del portafolio total invertida en deuda local de emisores no financieros.

**Interpretación económica u operativa**

Permite identificar la exposición del portafolio a deuda local de emisores no financieros y evaluar la diversificación de inversiones.

**Fórmula conceptual**

$$
\text{Deuda no financiera local} = \frac{\text{Valor invertido en deuda local de emisores no financieros}}{\text{Valor total del portafolio}} \times 100
$$

**Fórmula implementada**

$$
\text{Fila 33} = \text{LIMITES!AIOS!G4}
$$

El valor proviene del archivo `LIMITES`, hoja `AIOS`, o del campo mensual equivalente para esta categoría.

#### Fila 34: Acciones locales (%)

**¿Qué representa la fila 34?**

La fila 34 corresponde a la proporción del portafolio total invertida en acciones o renta variable local.

**Interpretación económica u operativa**

Permite identificar la exposición del portafolio a acciones o renta variable local y evaluar la diversificación de inversiones.

**Fórmula conceptual**

$$
\text{Acciones locales} = \frac{\text{Valor invertido en acciones o renta variable local}}{\text{Valor total del portafolio}} \times 100
$$

**Fórmula implementada**

$$
\text{Fila 34} = \text{LIMITES!AIOS!I4}
$$

El valor proviene del archivo `LIMITES`, hoja `AIOS`, o del campo mensual equivalente para esta categoría.

#### Fila 35: Fondos locales (%)

**¿Qué representa la fila 35?**

La fila 35 corresponde a la proporción del portafolio total invertida en fondos o vehículos colectivos locales.

**Interpretación económica u operativa**

Permite identificar la exposición del portafolio a fondos o vehículos colectivos locales y evaluar la diversificación de inversiones.

**Fórmula conceptual**

$$
\text{Fondos locales} = \frac{\text{Valor invertido en fondos o vehículos colectivos locales}}{\text{Valor total del portafolio}} \times 100
$$

**Fórmula implementada**

$$
\text{Fila 35} = \text{LIMITES!AIOS!K4}
$$

El valor proviene del archivo `LIMITES`, hoja `AIOS`, o del campo mensual equivalente para esta categoría.

#### Fila 36: Categoría local no usada

**¿Qué representa la fila 36?**

La fila 36 es una categoría reservada de la plantilla sin dato activo.

**Interpretación económica u operativa**

Mantiene la estructura de la plantilla y evita desplazar filas posteriores.

**Fórmula conceptual**

$$
\text{Fila 36} = 0
$$

**Fórmula implementada**

$$
\text{Fila 36} = 0
$$

Es una constante.

#### Fila 37: Deuda gubernamental exterior (%)

**¿Qué representa la fila 37?**

La fila 37 corresponde a la proporción del portafolio total invertida en deuda gubernamental extranjera.

**Interpretación económica u operativa**

Permite identificar la exposición del portafolio a deuda gubernamental extranjera y evaluar la diversificación de inversiones.

**Fórmula conceptual**

$$
\text{Deuda gubernamental exterior} = \frac{\text{Valor invertido en deuda gubernamental extranjera}}{\text{Valor total del portafolio}} \times 100
$$

**Fórmula implementada**

$$
\text{Fila 37} = \text{LIMITES!AIOS!dato exterior correspondiente}
$$

El valor proviene del archivo `LIMITES`, hoja `AIOS`, o del campo mensual equivalente para esta categoría.

#### Fila 38: Depósitos y efectivo exterior (%)

**¿Qué representa la fila 38?**

La fila 38 corresponde a la proporción del portafolio total invertida en liquidez o efectivo del exterior.

**Interpretación económica u operativa**

Permite identificar la exposición del portafolio a liquidez o efectivo del exterior y evaluar la diversificación de inversiones.

**Fórmula conceptual**

$$
\text{Depósitos y efectivo exterior} = \frac{\text{Valor invertido en liquidez o efectivo del exterior}}{\text{Valor total del portafolio}} \times 100
$$

**Fórmula implementada**

$$
\text{Fila 38} = \text{LIMITES!AIOS!dato exterior correspondiente}
$$

El valor proviene del archivo `LIMITES`, hoja `AIOS`, o del campo mensual equivalente para esta categoría.

#### Fila 39: Deuda no financiera exterior (%)

**¿Qué representa la fila 39?**

La fila 39 corresponde a la proporción del portafolio total invertida en deuda privada o no financiera del exterior.

**Interpretación económica u operativa**

Permite identificar la exposición del portafolio a deuda privada o no financiera del exterior y evaluar la diversificación de inversiones.

**Fórmula conceptual**

$$
\text{Deuda no financiera exterior} = \frac{\text{Valor invertido en deuda privada o no financiera del exterior}}{\text{Valor total del portafolio}} \times 100
$$

**Fórmula implementada**

$$
\text{Fila 39} = \text{LIMITES!AIOS!dato exterior correspondiente}
$$

El valor proviene del archivo `LIMITES`, hoja `AIOS`, o del campo mensual equivalente para esta categoría.

#### Fila 40: Acciones exterior (%)

**¿Qué representa la fila 40?**

La fila 40 corresponde a la proporción del portafolio total invertida en acciones o renta variable extranjera.

**Interpretación económica u operativa**

Permite identificar la exposición del portafolio a acciones o renta variable extranjera y evaluar la diversificación de inversiones.

**Fórmula conceptual**

$$
\text{Acciones exterior} = \frac{\text{Valor invertido en acciones o renta variable extranjera}}{\text{Valor total del portafolio}} \times 100
$$

**Fórmula implementada**

$$
\text{Fila 40} = \text{LIMITES!AIOS!dato exterior correspondiente}
$$

El valor proviene del archivo `LIMITES`, hoja `AIOS`, o del campo mensual equivalente para esta categoría.

#### Fila 41: Fondos exterior (%)

**¿Qué representa la fila 41?**

La fila 41 corresponde a la proporción del portafolio total invertida en fondos o vehículos colectivos del exterior.

**Interpretación económica u operativa**

Permite identificar la exposición del portafolio a fondos o vehículos colectivos del exterior y evaluar la diversificación de inversiones.

**Fórmula conceptual**

$$
\text{Fondos exterior} = \frac{\text{Valor invertido en fondos o vehículos colectivos del exterior}}{\text{Valor total del portafolio}} \times 100
$$

**Fórmula implementada**

$$
\text{Fila 41} = \text{LIMITES!AIOS!dato exterior correspondiente}
$$

El valor proviene del archivo `LIMITES`, hoja `AIOS`, o del campo mensual equivalente para esta categoría.

#### Fila 42: Referencia fija

**¿Qué representa la fila 42?**

La fila 42 es un valor fijo de referencia de la plantilla.

**Interpretación económica u operativa**

Funciona como marcador o control estructural del reporte.

**Fórmula conceptual**

$$
\text{Fila 42} = 2
$$

**Fórmula implementada**

$$
\text{Fila 42} = 2
$$

Es una constante.

#### Fila 43: Otros activos (%)

**¿Qué representa la fila 43?**

La fila 43 corresponde a la proporción del portafolio clasificada como otros activos.

**Interpretación económica u operativa**

Completa la clasificación de activos cuando existen rubros no incluidos en categorías principales.

**Fórmula conceptual**

$$
\text{Otros activos} = \frac{\text{Valor de otros activos}}{\text{Portafolio total}} \times 100
$$

**Fórmula implementada**

$$
\text{Fila 43} = \text{mensual.otros}
$$

Proviene de la lectura mensual de límites o composición.

#### Fila 44: Exposición exterior seleccionada (%)

**¿Qué representa la fila 44?**

La fila 44 corresponde a la suma de rubros exteriores seleccionados del archivo de límites.

**Interpretación económica u operativa**

Resume la exposición internacional agrupada en categorías específicas del portafolio.

**Fórmula conceptual**

$$
\text{Exposición exterior seleccionada} = \frac{\text{Rubros exteriores seleccionados}}{\text{Portafolio total}} \times 100
$$

**Fórmula implementada**

$$
\text{Fila 44} = (O4 + Q4 + S4 + U4 + W4 + Y4) \times 100
$$

Se calcula desde `LIMITES`, hoja `AIOS`.

#### Fila 45: Fondos / deuda gubernamental total (%)

**¿Qué representa la fila 45?**

La fila 45 corresponde a la relación entre fondos administrados y deuda gubernamental total.

**Interpretación económica u operativa**

Permite comparar el tamaño de los fondos de pensiones frente al saldo de deuda pública total.

**Fórmula conceptual**

$$
\text{Fondos sobre deuda pública total} = \frac{\text{Fondos administrados}}{\text{Deuda gubernamental total}} \times 100
$$

**Fórmula implementada**

$$
\text{Fila 45} = \frac{\text{Fila 28}}{\text{Deuda gubernamental total en USD}}
$$

La deuda gubernamental total proviene de `PIB_PEA_TRM_DG`, hoja `Hoja1`, columna `M`.

#### Fila 46: Referencia fija

**¿Qué representa la fila 46?**

La fila 46 es un valor fijo de la plantilla.

**Interpretación económica u operativa**

Funciona como control estructural del reporte.

**Fórmula conceptual**

$$
\text{Fila 46} = 4
$$

**Fórmula implementada**

$$
\text{Fila 46} = 4
$$

Es una constante.

#### Fila 47: Participación Protección + Porvenir (%)

**¿Qué representa la fila 47?**

La fila 47 corresponde a la participación conjunta de Protección y Porvenir sobre el total del sistema. A diferencia de la columna Q mensual, esta fila no mide concentración de afiliados/personas sino concentración sobre saldos/fondos administrados según la fuente `SISTEMA TOTAL`.

**Interpretación económica u operativa**

Mide concentración de mercado de las administradoras seleccionadas dentro del total de fondos.

**Fórmula conceptual**

$$
\text{Participación Protección y Porvenir} = \frac{\text{Fondos Protección} + \text{Fondos Porvenir}}{\text{Fondos totales del sistema}} \times 100
$$

**Fórmula implementada**

$$
\text{Fila 47} = \frac{\text{restot!C14} + \text{restot!D14}}{\text{restot!J14}}
$$

Se lee de `SISTEMA TOTAL`, hoja `restot`.

#### Fila 48: Activos en USD

**¿Qué representa la fila 48?**

La fila 48 corresponde a activos contables convertidos a dólares.

**Interpretación económica u operativa**

Mide el tamaño del balance por el lado de activos en una moneda comparable.

**Fórmula conceptual**

$$
\text{Activos en USD} = \frac{\text{Activos en COP}}{\text{TRM}}
$$

**Fórmula implementada**

$$
\text{Fila 48} = \frac{\text{mensual.activosCuentas}}{\text{TRM}}
$$

La fuente técnica exacta está descrita en la tabla de fórmulas semestrales.

#### Fila 49: Pasivos en USD

**¿Qué representa la fila 49?**

La fila 49 corresponde a pasivos contables convertidos a dólares.

**Interpretación económica u operativa**

Mide las obligaciones del balance en una moneda comparable.

**Fórmula conceptual**

$$
\text{Pasivos en USD} = \frac{\text{Pasivos en COP}}{\text{TRM}}
$$

**Fórmula implementada**

$$
\text{Fila 49} = \frac{\text{mensual.pasivosCuentas}}{\text{TRM}}
$$

La fuente técnica exacta está descrita en la tabla de fórmulas semestrales.

#### Fila 50: Patrimonio en USD

**¿Qué representa la fila 50?**

La fila 50 corresponde a patrimonio contable convertido a dólares.

**Interpretación económica u operativa**

Mide el valor patrimonial contable de la entidad o sistema reportado.

**Fórmula conceptual**

$$
\text{Patrimonio en USD} = \frac{\text{Activos} - \text{Pasivos}}{\text{TRM}}
$$

**Fórmula implementada**

$$
\text{Fila 50} = \frac{\text{Activos cuentas} - \text{Pasivos cuentas}}{\text{TRM}}
$$

La fuente técnica exacta está descrita en la tabla de fórmulas semestrales.

#### Fila 51: Comisiones

**¿Qué representa la fila 51?**

La fila 51 corresponde a ingresos por comisiones según la plantilla contable.

**Interpretación económica u operativa**

Mide ingresos operacionales asociados a la administración de recursos.

**Fórmula conceptual**

$$
\text{Comisiones} = \text{Ingresos por comisiones reportados}
$$

**Fórmula implementada**

$$
\text{Fila 51} = \text{CUENTAS!E13}
$$

La fuente técnica exacta está descrita en la tabla de fórmulas semestrales.

#### Fila 52: Gastos

**¿Qué representa la fila 52?**

La fila 52 corresponde a gastos reportados en la plantilla contable.

**Interpretación económica u operativa**

Mide egresos operativos o administrativos del periodo.

**Fórmula conceptual**

$$
\text{Gastos} = \text{Gastos reportados}
$$

**Fórmula implementada**

$$
\text{Fila 52} = \text{CUENTAS!G15}
$$

La fuente técnica exacta está descrita en la tabla de fórmulas semestrales.

#### Fila 53: Resultado operacional

**¿Qué representa la fila 53?**

La fila 53 corresponde a resultado de operación antes del resultado neto.

**Interpretación económica u operativa**

Mide desempeño operativo de la entidad o sistema reportado.

**Fórmula conceptual**

$$
\text{Resultado operacional} = \text{Ingresos operacionales} - \text{Gastos operacionales}
$$

**Fórmula implementada**

$$
\text{Fila 53} = \text{CUENTAS!E41}
$$

La fuente técnica exacta está descrita en la tabla de fórmulas semestrales.

#### Fila 54: Resultado neto

**¿Qué representa la fila 54?**

La fila 54 corresponde a resultado final del periodo.

**Interpretación económica u operativa**

Mide utilidad o pérdida final después de ingresos y gastos.

**Fórmula conceptual**

$$
\text{Resultado neto} = \text{Ingresos totales} - \text{Gastos totales}
$$

**Fórmula implementada**

$$
\text{Fila 54} = \text{CUENTAS!E44}
$$

La fuente técnica exacta está descrita en la tabla de fórmulas semestrales.

#### Fila 55: Gastos de administración

**¿Qué representa la fila 55?**

La fila 55 corresponde a rubro administrativo seleccionado.

**Interpretación económica u operativa**

Mide la carga administrativa relevante para los indicadores del reporte.

**Fórmula conceptual**

$$
\text{Administración} = \text{Gasto administrativo reportado}
$$

**Fórmula implementada**

$$
\text{Fila 55} = \text{CUENTAS!H24}
$$

La fuente técnica exacta está descrita en la tabla de fórmulas semestrales.

#### Fila 56: Cuenta 511500 en USD

**¿Qué representa la fila 56?**

La fila 56 corresponde a cuenta 511500 convertida a dólares.

**Interpretación económica u operativa**

Permite analizar este rubro específico en moneda comparable.

**Fórmula conceptual**

$$
\text{Cuenta 511500 en USD} = \frac{\text{Cuenta 511500 en COP}}{\text{TRM}}
$$

**Fórmula implementada**

$$
\text{Fila 56} = \frac{\text{CUENTAS!C21}}{\text{TRM}}
$$

La fuente técnica exacta está descrita en la tabla de fórmulas semestrales.

#### Fila 57: Cuenta 511527 en USD

**¿Qué representa la fila 57?**

La fila 57 corresponde a cuenta 511527 convertida a dólares.

**Interpretación económica u operativa**

Permite analizar afiliaciones a fondos de pensiones en moneda comparable.

**Fórmula conceptual**

$$
\text{Cuenta 511527 en USD} = \frac{\text{Cuenta 511527 en COP}}{\text{TRM}}
$$

**Fórmula implementada**

$$
\text{Fila 57} = \frac{\text{CUENTAS!C22}}{\text{TRM}}
$$

La fuente técnica exacta está descrita en la tabla de fórmulas semestrales.

#### Fila 58: Cuentas 511500 + 511527 en USD

**¿Qué representa la fila 58?**

La fila 58 corresponde a suma de las cuentas 511500 y 511527 convertida a dólares.

**Interpretación económica u operativa**

Resume los rubros de comisión y afiliación solicitados para el análisis.

**Fórmula conceptual**

$$
\text{Cuentas 511500 y 511527 en USD} = \frac{\text{Cuenta 511500} + \text{Cuenta 511527}}{\text{TRM}}
$$

**Fórmula implementada**

$$
\text{Fila 58} = \frac{\text{CUENTAS!C21} + \text{CUENTAS!C22}}{\text{TRM}}
$$

La fuente técnica exacta está descrita en la tabla de fórmulas semestrales.

#### Fila 59: Otros gastos operacionales en USD

**¿Qué representa la fila 59?**

La fila 59 corresponde a suma de cuentas seleccionadas de gastos operacionales convertida a dólares.

**Interpretación económica u operativa**

Mide la carga de gastos operacionales distintos de los rubros principales.

**Fórmula conceptual**

$$
\text{Otros gastos operacionales en USD} = \frac{\text{Suma de cuentas de gastos seleccionadas}}{\text{TRM}}
$$

**Fórmula implementada**

$$
\text{Fila 59} = \frac{C24+C28+C29+C31+C32+C33+C34+C35+C36+C37+C38}{\text{TRM}}
$$

La fuente técnica exacta está descrita en la tabla de fórmulas semestrales.

#### Fila 60: Gasto de operación 510000 en USD

**¿Qué representa la fila 60?**

La fila 60 corresponde a cuenta 510000 convertida a dólares.

**Interpretación económica u operativa**

Mide el gasto operacional agregado en moneda comparable.

**Fórmula conceptual**

$$
\text{Gasto operación 510000 en USD} = \frac{\text{Cuenta 510000}}{\text{TRM}}
$$

**Fórmula implementada**

$$
\text{Fila 60} = \frac{\text{CUENTAS!C15}}{\text{TRM}}
$$

La fuente técnica exacta está descrita en la tabla de fórmulas semestrales.

#### Fila 61: Aportes recibidos por aportante

**¿Qué representa la fila 61?**

La fila 61 corresponde a aportes recibidos convertidos a dólares y normalizados por aportantes en miles.

**Interpretación económica u operativa**

Mide la intensidad de aportes recibidos por población aportante.

**Fórmula conceptual**

$$
\text{Aportes por aportante} = \frac{\text{Aportes recibidos en USD}}{\text{Aportantes}/1000} \times 1000
$$

**Fórmula implementada**

$$
\text{Fila 61} = \frac{\text{Formato 136!G6}/\text{TRM}}{\text{Aportantes}/1000} \times 1000
$$

La fuente técnica exacta está descrita en la tabla de fórmulas semestrales.

#### Fila 62: Gastos / aportes recibidos (%)

**¿Qué representa la fila 62?**

La fila 62 corresponde a gastos frente a aportes recibidos.

**Interpretación económica u operativa**

Mide la carga de gastos sobre el flujo de aportes recibidos.

**Fórmula conceptual**

$$
\text{Gastos sobre aportes} = \frac{\text{Gastos}}{\text{Aportes recibidos}} \times 100
$$

**Fórmula implementada**

$$
\text{Fila 62} = \frac{\text{cuentas.gastos}}{\text{aportesRecibidos}/\text{TRM}} \times 100
$$

La fuente técnica exacta está descrita en la tabla de fórmulas semestrales.

#### Fila 63: Patrimonio / fondos administrados (%)

**¿Qué representa la fila 63?**

La fila 63 corresponde a patrimonio frente al valor de fondos administrados.

**Interpretación económica u operativa**

Mide respaldo patrimonial relativo al tamaño de recursos administrados.

**Fórmula conceptual**

$$
\text{Patrimonio sobre fondos} = \frac{\text{Patrimonio}}{\text{Fondos administrados}} \times 100
$$

**Fórmula implementada**

$$
\text{Fila 63} = \frac{\text{Patrimonio base mes}/\text{TRM}}{\text{Fila 28}} \times 100
$$

La fuente técnica exacta está descrita en la tabla de fórmulas semestrales.

#### Fila 64: Patrimonio por afiliado

**¿Qué representa la fila 64?**

La fila 64 corresponde a patrimonio distribuido por afiliado.

**Interpretación económica u operativa**

Mide respaldo patrimonial promedio asociado a cada afiliado.

**Fórmula conceptual**

$$
\text{Patrimonio por afiliado} = \frac{\text{Patrimonio}}{\text{Afiliados}} \times 1{,}000{,}000
$$

**Fórmula implementada**

$$
\text{Fila 64} = \frac{\text{Fila 50}}{\text{Afiliados}} \times 1{,}000{,}000
$$

La fuente técnica exacta está descrita en la tabla de fórmulas semestrales.

#### Fila 65: Resultado neto / comisiones (%)

**¿Qué representa la fila 65?**

La fila 65 corresponde a resultado neto frente a comisiones.

**Interpretación económica u operativa**

Mide margen o rentabilidad sobre ingresos por comisiones.

**Fórmula conceptual**

$$
\text{Resultado sobre comisiones} = \frac{\text{Resultado neto}}{\text{Comisiones}} \times 100
$$

**Fórmula implementada**

$$
\text{Fila 65} = \frac{\text{Fila 54}}{\text{Fila 51}} \times 100
$$

La fuente técnica exacta está descrita en la tabla de fórmulas semestrales.

#### Fila 66: Resultado neto / patrimonio (%)

**¿Qué representa la fila 66?**

La fila 66 corresponde a resultado neto frente al patrimonio.

**Interpretación económica u operativa**

Aproxima la rentabilidad sobre patrimonio.

**Fórmula conceptual**

$$
\text{Resultado sobre patrimonio} = \frac{\text{Resultado neto}}{\text{Patrimonio}} \times 100
$$

**Fórmula implementada**

$$
\text{Fila 66} = \frac{\text{Fila 54}}{\text{Fila 50}} \times 100
$$

La fuente técnica exacta está descrita en la tabla de fórmulas semestrales.

#### Fila 67: Gastos por afiliado

**¿Qué representa la fila 67?**

La fila 67 corresponde a gastos distribuidos por afiliado.

**Interpretación económica u operativa**

Mide costo promedio asociado a cada afiliado.

**Fórmula conceptual**

$$
\text{Gasto por afiliado} = \frac{\text{Gastos}}{\text{Afiliados}} \times 1{,}000{,}000
$$

**Fórmula implementada**

$$
\text{Fila 67} = \frac{\text{Fila 52}}{\text{Afiliados}} \times 1{,}000{,}000
$$

La fuente técnica exacta está descrita en la tabla de fórmulas semestrales.

#### Fila 68: Comisiones por aportante

**¿Qué representa la fila 68?**

La fila 68 corresponde a comisiones distribuidas por aportante.

**Interpretación económica u operativa**

Mide ingreso promedio por aportante.

**Fórmula conceptual**

$$
\text{Comisión por aportante} = \frac{\text{Comisiones}}{\text{Aportantes}} \times 1{,}000{,}000
$$

**Fórmula implementada**

$$
\text{Fila 68} = \frac{\text{Fila 51}}{\text{Aportantes}} \times 1{,}000{,}000
$$

La fuente técnica exacta está descrita en la tabla de fórmulas semestrales.

#### Fila 69: Administración / aportes por aportante

**¿Qué representa la fila 69?**

La fila 69 corresponde a gasto administrativo frente al indicador de aportes por aportante.

**Interpretación económica u operativa**

Mide carga administrativa relativa al flujo de aportes normalizado.

**Fórmula conceptual**

$$
\text{Administración sobre aportes por aportante} = \frac{\text{Administración}}{\text{Aportes por aportante}}
$$

**Fórmula implementada**

$$
\text{Fila 69} = \frac{\text{Fila 55}}{\text{Fila 61}}
$$

La fuente técnica exacta está descrita en la tabla de fórmulas semestrales.

#### Fila 70: Referencia fija

**¿Qué representa la fila 70?**

La fila 70 corresponde a valor fijo de referencia.

**Interpretación económica u operativa**

Mantiene la estructura de la plantilla.

**Fórmula conceptual**

$$
\text{Fila 70} = 16
$$

**Fórmula implementada**

$$
\text{Fila 70} = 16
$$

La fuente técnica exacta está descrita en la tabla de fórmulas semestrales.

#### Fila 71: Comisión promedio obligatoria (%)

**¿Qué representa la fila 71?**

La fila 71 corresponde a promedio de comisiones obligatorias de administradoras seleccionadas.

**Interpretación económica u operativa**

Mide el costo promedio de comisión obligatoria.

**Fórmula conceptual**

$$
\text{Comisión promedio} = \frac{\text{COL} + \text{POR} + \text{PRO} + \text{SKA}}{4} \times 100
$$

**Fórmula implementada**

$$
\text{Fila 71} = \text{promedio(comisiones obligatorias)} \times 100
$$

La fuente técnica exacta está descrita en la tabla de fórmulas semestrales.

#### Fila 72: Referencia cero

**¿Qué representa la fila 72?**

La fila 72 corresponde a rubro sin dato activo.

**Interpretación económica u operativa**

Reserva de plantilla.

**Fórmula conceptual**

$$
\text{Fila 72} = 0
$$

**Fórmula implementada**

$$
\text{Fila 72} = 0
$$

La fuente técnica exacta está descrita en la tabla de fórmulas semestrales.

#### Fila 73: Referencia cero

**¿Qué representa la fila 73?**

La fila 73 corresponde a rubro sin dato activo.

**Interpretación económica u operativa**

Reserva de plantilla.

**Fórmula conceptual**

$$
\text{Fila 73} = 0
$$

**Fórmula implementada**

$$
\text{Fila 73} = 0
$$

La fuente técnica exacta está descrita en la tabla de fórmulas semestrales.

#### Fila 74: Aporte trabajador

**¿Qué representa la fila 74?**

La fila 74 corresponde a parte trabajador de la diferencia entre 3 y comisión promedio.

**Interpretación económica u operativa**

Estima la distribución del aporte residual hacia trabajador.

**Fórmula conceptual**

$$
\text{Aporte trabajador} = (3 - \text{Comisión promedio}) \times 0.25
$$

**Fórmula implementada**

$$
\text{Fila 74} = (3 - \text{Fila 71}) \times 0.25
$$

La fuente técnica exacta está descrita en la tabla de fórmulas semestrales.

#### Fila 75: Aporte empleador

**¿Qué representa la fila 75?**

La fila 75 corresponde a parte empleador de la diferencia entre 3 y comisión promedio.

**Interpretación económica u operativa**

Estima la distribución del aporte residual hacia empleador.

**Fórmula conceptual**

$$
\text{Aporte empleador} = (3 - \text{Comisión promedio}) \times 0.75
$$

**Fórmula implementada**

$$
\text{Fila 75} = (3 - \text{Fila 71}) \times 0.75
$$

La fuente técnica exacta está descrita en la tabla de fórmulas semestrales.

#### Fila 76: Referencia cero

**¿Qué representa la fila 76?**

La fila 76 corresponde a rubro sin dato activo.

**Interpretación económica u operativa**

Reserva de plantilla.

**Fórmula conceptual**

$$
\text{Fila 76} = 0
$$

**Fórmula implementada**

$$
\text{Fila 76} = 0
$$

La fuente técnica exacta está descrita en la tabla de fórmulas semestrales.

#### Fila 77: Comisiones

**¿Qué representa la fila 77?**

La fila 77 corresponde a comisiones contables reutilizadas.

**Interpretación económica u operativa**

Base para medir comisiones respecto a fondos.

**Fórmula conceptual**

$$
\text{Comisiones} = \text{Comisiones reportadas}
$$

**Fórmula implementada**

$$
\text{Fila 77} = \text{Fila 51}
$$

La fuente técnica exacta está descrita en la tabla de fórmulas semestrales.

#### Fila 78: Fondos administrados

**¿Qué representa la fila 78?**

La fila 78 corresponde a valor de fondos administrados reutilizado.

**Interpretación económica u operativa**

Base de activos administrados para indicadores de eficiencia.

**Fórmula conceptual**

$$
\text{Fondos administrados} = \text{Valor total de fondos}
$$

**Fórmula implementada**

$$
\text{Fila 78} = \text{Fila 28}
$$

La fuente técnica exacta está descrita en la tabla de fórmulas semestrales.

#### Fila 79: Comisiones / fondos

**¿Qué representa la fila 79?**

La fila 79 corresponde a comisiones frente a fondos administrados.

**Interpretación económica u operativa**

Mide peso de comisiones sobre activos administrados.

**Fórmula conceptual**

$$
\text{Comisiones sobre fondos} = \frac{\text{Comisiones}}{\text{Fondos administrados}} \times 100
$$

**Fórmula implementada**

$$
\text{Fila 79} = \frac{\text{Fila 77}}{\text{Fila 78}}
$$

La fuente técnica exacta está descrita en la tabla de fórmulas semestrales.

#### Fila 80: Años desde 1994

**¿Qué representa la fila 80?**

La fila 80 corresponde a años transcurridos desde el año base 1994.

**Interpretación económica u operativa**

Mide antigüedad del periodo de referencia.

**Fórmula conceptual**

$$
\text{Años desde 1994} = \text{Año de corte} - 1994
$$

**Fórmula implementada**

$$
\text{Fila 80} = \text{fechaCorte.year} - 1994
$$

La fuente técnica exacta está descrita en la tabla de fórmulas semestrales.

#### Fila 81: Sin información disponible

**¿Qué representa la fila 81?**

No se tiene información disponible para la fila 81 del archivo semestral dentro del mapeo implementado actualmente.

**Interpretación económica u operativa**

No es posible asignar una interpretación económica u operativa sin una fuente o definición funcional confirmada para esta fila.

**Fórmula conceptual**

No aplica.

**Fórmula implementada**

No aplica.

No se encontró escritura de la fila 81 en `SemestralExcelGenerator`; por tanto se documenta explícitamente como información no disponible.

#### Fila 82: Rentabilidad nominal 10 años

**¿Qué representa la fila 82?**

La fila 82 corresponde al retorno nominal anualizado para el horizonte de 10 años.

**Interpretación económica u operativa**

Mide desempeño financiero a 10 años sin descontar inflación.

**Fórmula conceptual**

$$
\text{Rentabilidad nominal} = \left(\left(\frac{\text{NAV final}}{\text{NAV inicial}}\right)^{1/n} - 1\right) \times 100
$$

**Fórmula implementada**

$$
\text{Fila 82} = \text{RentabilidadService(10 años, nominal)}
$$

Se calcula con NAV de `Valores_Fondo_Moder` e IPC de `Rent_Vr_Uni_Moderado` según corresponda.

#### Fila 83: Rentabilidad real 10 años

**¿Qué representa la fila 83?**

La fila 83 corresponde al retorno real anualizado para el horizonte de 10 años.

**Interpretación económica u operativa**

Mide desempeño financiero a 10 años descontando inflación.

**Fórmula conceptual**

$$
\text{Rentabilidad real} = \left(\frac{1+\text{Rentabilidad nominal}}{1+\text{Inflación anualizada}} - 1\right) \times 100
$$

**Fórmula implementada**

$$
\text{Fila 83} = \text{RentabilidadService(10 años, real)}
$$

Se calcula con NAV de `Valores_Fondo_Moder` e IPC de `Rent_Vr_Uni_Moderado` según corresponda.

#### Fila 84: Rentabilidad nominal 5 años

**¿Qué representa la fila 84?**

La fila 84 corresponde al retorno nominal anualizado para el horizonte de 5 años.

**Interpretación económica u operativa**

Mide desempeño financiero a 5 años sin descontar inflación.

**Fórmula conceptual**

$$
\text{Rentabilidad nominal} = \left(\left(\frac{\text{NAV final}}{\text{NAV inicial}}\right)^{1/n} - 1\right) \times 100
$$

**Fórmula implementada**

$$
\text{Fila 84} = \text{RentabilidadService(5 años, nominal)}
$$

Se calcula con NAV de `Valores_Fondo_Moder` e IPC de `Rent_Vr_Uni_Moderado` según corresponda.

#### Fila 85: Rentabilidad real 5 años

**¿Qué representa la fila 85?**

La fila 85 corresponde al retorno real anualizado para el horizonte de 5 años.

**Interpretación económica u operativa**

Mide desempeño financiero a 5 años descontando inflación.

**Fórmula conceptual**

$$
\text{Rentabilidad real} = \left(\frac{1+\text{Rentabilidad nominal}}{1+\text{Inflación anualizada}} - 1\right) \times 100
$$

**Fórmula implementada**

$$
\text{Fila 85} = \text{RentabilidadService(5 años, real)}
$$

Se calcula con NAV de `Valores_Fondo_Moder` e IPC de `Rent_Vr_Uni_Moderado` según corresponda.

#### Fila 86: Rentabilidad nominal 3 años

**¿Qué representa la fila 86?**

La fila 86 corresponde al retorno nominal anualizado para el horizonte de 3 años.

**Interpretación económica u operativa**

Mide desempeño financiero a 3 años sin descontar inflación.

**Fórmula conceptual**

$$
\text{Rentabilidad nominal} = \left(\left(\frac{\text{NAV final}}{\text{NAV inicial}}\right)^{1/n} - 1\right) \times 100
$$

**Fórmula implementada**

$$
\text{Fila 86} = \text{RentabilidadService(3 años, nominal)}
$$

Se calcula con NAV de `Valores_Fondo_Moder` e IPC de `Rent_Vr_Uni_Moderado` según corresponda.

#### Fila 87: Rentabilidad real 3 años

**¿Qué representa la fila 87?**

La fila 87 corresponde al retorno real anualizado para el horizonte de 3 años.

**Interpretación económica u operativa**

Mide desempeño financiero a 3 años descontando inflación.

**Fórmula conceptual**

$$
\text{Rentabilidad real} = \left(\frac{1+\text{Rentabilidad nominal}}{1+\text{Inflación anualizada}} - 1\right) \times 100
$$

**Fórmula implementada**

$$
\text{Fila 87} = \text{RentabilidadService(3 años, real)}
$$

Se calcula con NAV de `Valores_Fondo_Moder` e IPC de `Rent_Vr_Uni_Moderado` según corresponda.

#### Fila 88: Rentabilidad nominal 1 año

**¿Qué representa la fila 88?**

La fila 88 corresponde al retorno nominal anualizado para el horizonte de 1 año.

**Interpretación económica u operativa**

Mide desempeño financiero a 1 año sin descontar inflación.

**Fórmula conceptual**

$$
\text{Rentabilidad nominal} = \left(\left(\frac{\text{NAV final}}{\text{NAV inicial}}\right)^{1/n} - 1\right) \times 100
$$

**Fórmula implementada**

$$
\text{Fila 88} = \text{RentabilidadService(1 año, nominal)}
$$

Se calcula con NAV de `Valores_Fondo_Moder` e IPC de `Rent_Vr_Uni_Moderado` según corresponda.

#### Fila 89: Rentabilidad real 1 año

**¿Qué representa la fila 89?**

La fila 89 corresponde al retorno real anualizado para el horizonte de 1 año.

**Interpretación económica u operativa**

Mide desempeño financiero a 1 año descontando inflación.

**Fórmula conceptual**

$$
\text{Rentabilidad real} = \left(\frac{1+\text{Rentabilidad nominal}}{1+\text{Inflación anualizada}} - 1\right) \times 100
$$

**Fórmula implementada**

$$
\text{Fila 89} = \text{RentabilidadService(1 año, real)}
$$

Se calcula con NAV de `Valores_Fondo_Moder` e IPC de `Rent_Vr_Uni_Moderado` según corresponda.

## 7. Parámetros de Formato 136 para fila 61

| Celda | Valor escrito antes de evaluar `G6` | Ejemplo con corte `30/06/2025` |
|---|---|---|
| `C7` | `fechaCorte.minusYears(1).withDayOfMonth(1)` | `01/06/2024` |
| `D6` | `fechaCorte` | `30/06/2025` |
| `D7` | `fechaCorte` | `30/06/2025` |
