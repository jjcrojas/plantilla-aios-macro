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

## 5. Parámetros de Formato 136 para fila 61

| Celda | Valor escrito antes de evaluar `G6` | Ejemplo con corte `30/06/2025` |
|---|---|---|
| `C7` | `fechaCorte.minusYears(1).withDayOfMonth(1)` | `01/06/2024` |
| `D6` | `fechaCorte` | `30/06/2025` |
| `D7` | `fechaCorte` | `30/06/2025` |
