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
| B | `afiliados = hombres + mujeres` | Formato 491, `informe de prensa`, `C11 + D11` | Valor crudo (personas) |
| C | `aportantes` | Formato 491, `multifondos`, `E25` | Valor crudo (personas) |
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
| Q | `consFdosAdmon` | Formato 491, `multifondos`, cálculo con `J8`, `J9`, `J12` | Valor crudo / ratio según plantilla |
| R | `porcVrFondo` | `SISTEMA TOTAL`, `restot`, `(Protección + Porvenir) / Sistema` | Ratio / porcentaje según plantilla |
| S | `TRM` | `PIB_PEA_TRM_DG`, último valor aplicable a la fecha de corte | COP por USD |

## 3. Informe trimestral

El informe trimestral escribe mapas por hoja. Las fórmulas exactas dependen de los mapas generados por `TrimestralDataReader`; la tabla resume la forma de escritura y unidad.

| Hoja | Fórmula / mapeo | Fuente principal | Unidad escrita |
|---|---|---|---|
| `afiliados` | Valores por fondo y administradora (`mod_*`, `con_*`, `mr_*`, combinaciones) | Formato 491 / archivos trimestrales de referencia | Valor crudo (personas) |
| `aportantes` | `colf`, `porv`, `prot`, `sk` y ceros para entidades no aplicables | Formato 491 / datos trimestrales | Valor crudo (personas) |
| `colombia` | Saldos por fondo/administradora, con agregados como `mod_sk + mod_alt` | `SISTEMA TOTAL` y datos de fondos | USD o MM USD según plantilla |
| `traspasos` | Traspasos por administradora | Formato 493 | Valor crudo |
| `gastos` | `gastoNetoCOP / TRM` | Base anual / cuentas de gasto; TRM de `PIB_PEA_TRM_DG` | USD |
| `promotores` | Ceros cuando no hay fuente disponible | Constante temporal | Valor crudo |
| `rentabilidad` | Rentabilidad nominal y real por administradora/fondo | `Rent_Vr_Uni_Moderado` y/o series de rentabilidad | Porcentaje |
| `comisiones` | Comisiones obligatorias por administradora (`col_obl`, `por_obl`, `pro_obl`, `ska_obl`) | Datos trimestrales de comisiones | Porcentaje |

## 4. Informe semestral (`semestral.xlsx`)

| Fila | Fórmula / dato | Fuente, hoja y celda | Unidad escrita |
|---:|---|---|---|
| 3 | `afiliados = hombres + mujeres` | Formato 491, `informe de prensa`, `C11 + D11` | Valor crudo (personas) |
| 4 | `(afiliadosMenor30 / afiliados) * 100` | Formato 491, `informe de prensa`, `C81 + D81`; fila 3 | Porcentaje |
| 5 | `(afiliados30a44 / afiliados) * 100` | Formato 491, `informe de prensa`, `C82 + D82`; fila 3 | Porcentaje |
| 6 | `(afiliados45a59 / afiliados) * 100` | Formato 491, `informe de prensa`, `C83 + D83`; fila 3 | Porcentaje |
| 7 | `(afiliadosMayor60 / afiliados) * 100` | Formato 491, `informe de prensa`, `C84 + D84`; fila 3 | Porcentaje |
| 8 | `100` | Constante | Porcentaje |
| 9 | `afiliados / 1000` | Formato 491; fila 3 | Miles de personas |
| 10 | `(mujeres / afiliados) * 100` | Formato 491, `informe de prensa`, `D11`; fila 3 | Porcentaje |
| 11 | `aportantes` | Formato 491, `multifondos`, `E25` | Valor crudo (personas) |
| 12 | `(afiliados / PEA) * 100` | Afiliados de Formato 491; PEA de `PIB_PEA_TRM_DG` | Porcentaje |
| 13 | `(aportantes / PEA) * 100` | Aportantes de Formato 491; PEA de `PIB_PEA_TRM_DG` | Porcentaje |
| 14 | `(aportantes / afiliados) * 100` | Formato 491 | Porcentaje |
| 15 | `salario mínimo Colombia COP / TRM` | Formato 491, hoja `SM COLOMBIA`, `E8`; TRM de `PIB_PEA_TRM_DG` | USD |
| 16 | `total pensionados` | Formato 495, `TOTAL PENSIONADOS`, parámetro `B4`, valor en columna `I` para la fecha | Valor crudo (personas) |
| 17 | `por Entidad!BI62 / fila16` | Formato 495, hoja `por Entidad`, parámetro `C6`, celda `BI62` | Ratio / porcentaje |
| 18 | `por Entidad!BH62 / fila16` | Formato 495, hoja `por Entidad`, parámetro `C6`, celda `BH62` | Ratio / porcentaje |
| 19 | `por Entidad!BJ62 / fila16` | Formato 495, hoja `por Entidad`, parámetro `C6`, celda `BJ62` | Ratio / porcentaje |
| 26 | `traspasosSistema` | Formato 493 / lector mensual | Valor crudo |
| 27 | `traspasosSistema / afiliados` | Formato 493 y Formato 491 | Porcentaje (formato Excel) |
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
| 61 | `(aportesRecibidos136 / TRM) / (aportantes / 1000) * 1000` | `Formato_136_Meses`, hoja `FORMATO OBL`, parámetros `C7`, `D6`, `D7`, resultado `G6`; aportantes de Formato 491; TRM | USD por mil aportantes |
| 62 | `gastos / (aportesRecibidos136 / TRM) * 100` | Gastos de `CUENTAS`; aportes de Formato 136; TRM | Porcentaje |
| 63 | `(patrimonioBaseMesMMCop / TRM) / fila28 * 100` | `Plantilla AIOS-probable`, base mes; TRM; fila 28 | Porcentaje |
| 64 | `patrimonioUsd / afiliados * 1,000,000` | Fila 50 y afiliados de Formato 491 | USD por afiliado |
| 65 | `resultadoNeto / comisiones * 100` | `CUENTAS`, filas 54 y 51 | Porcentaje |
| 66 | `resultadoNeto / patrimonioUsd * 100` | `CUENTAS` y fila 50 | Porcentaje |
| 67 | `gastos / afiliados * 1,000,000` | `CUENTAS` y Formato 491 | Valor por afiliado |
| 68 | `comisiones / aportantes * 1,000,000` | `CUENTAS` y Formato 491 | Valor por aportante |
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

### 5.1 Fila 29: Fondos administrados / PIB (%)

**¿Qué representa la fila 29?**

La fila 29 del archivo `Semestral_Colombia.xlsx` corresponde a la proporción del valor total de los fondos de pensiones respecto al Producto Interno Bruto (PIB) del país, expresada como porcentaje (%).

**Interpretación económica**

Este indicador responde: **¿qué tan grande es el sistema de fondos de pensiones frente al tamaño total de la economía?** Un valor mayor indica que los activos administrados por los fondos tienen un peso más alto frente a la producción anual del país.

**Fórmula conceptual**

$$
\text{Fondos / PIB (\%)} = \frac{\text{Valor total de los fondos de pensiones}}{\text{Producto Interno Bruto (PIB)}} \times 100
$$

**Fórmula implementada**

$$
\text{Fila 29} = \frac{\text{Fila 28}}{\text{PIB semestral en COP} / \text{TRM}}
$$

En la implementación, `fila 28` ya está expresada en millones de USD; el PIB se convierte con TRM para que ambas magnitudes sean comparables.

### 5.2 Fila 31: Deuda gubernamental local (%)

**¿Qué representa la fila 31?**

La fila 31 del archivo `Semestral_Colombia.xlsx` corresponde a la proporción del portafolio total de los fondos de pensiones que está invertida en deuda gubernamental local o interna, expresada como porcentaje (%) del total del portafolio.

**Interpretación económica**

Este indicador muestra la exposición del portafolio administrado a títulos de deuda pública interna. Permite evaluar qué parte de los recursos pensionales financia deuda gubernamental local y qué tan concentrada está la inversión en este tipo de activo.

**Fórmula conceptual**

$$
\text{Deuda gubernamental local (\%)} = \frac{\text{Valor invertido en deuda gubernamental local}}{\text{Valor total del portafolio de fondos}} \times 100
$$

**Fórmula implementada**

$$
\text{Fila 31} = \text{LIMITES!AIOS!C4}
$$

El archivo `LIMITES` ya entrega la proporción calculada para la categoría de deuda gubernamental local. Por eso la fila toma directamente el ratio de la celda `C4`.

### 5.3 Otras fórmulas conceptuales clave

$$
\text{Afiliados / PEA (\%)} = \frac{\text{Afiliados}}{\text{Población Económicamente Activa}} \times 100
$$

$$
\text{Aportantes / PEA (\%)} = \frac{\text{Aportantes}}{\text{Población Económicamente Activa}} \times 100
$$

$$
\text{Traspasos / Afiliados (\%)} = \frac{\text{Traspasos del sistema}}{\text{Afiliados}} \times 100
$$

$$
\text{Fondos / Deuda pública total (\%)} = \frac{\text{Fondos administrados}}{\text{Deuda gubernamental total}} \times 100
$$

$$
\text{Participación Protección + Porvenir (\%)} = \frac{\text{Fondos de Protección} + \text{Fondos de Porvenir}}{\text{Fondos totales del sistema}} \times 100
$$

$$
\text{Resultado neto / Patrimonio (\%)} = \frac{\text{Resultado neto}}{\text{Patrimonio}} \times 100
$$

$$
\text{Comisiones / Fondos (\%)} = \frac{\text{Comisiones}}{\text{Fondos administrados}} \times 100
$$

## 6. Parámetros de Formato 136 para fila 61

| Celda | Valor escrito antes de evaluar `G6` | Ejemplo con corte `30/06/2025` |
|---|---|---|
| `C7` | `fechaCorte.minusYears(1).withDayOfMonth(1)` | `01/06/2024` |
| `D6` | `fechaCorte` | `30/06/2025` |
| `D7` | `fechaCorte` | `30/06/2025` |
