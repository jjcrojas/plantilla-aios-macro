# Generación de datos de los informes AIOS

## 1. Propósito y forma de leer este documento

Este documento explica cómo se obtiene cada dato de los tres archivos AIOS que genera la aplicación:

- `Boletin_AIOS MENSUAL.xlsx`: una fila por mes.
- `Boletin_AIOS TRIMESTRAL.xlsx`: una fila por cierre trimestral en cada hoja temática.
- `semestral.xlsx`: una columna por cierre de junio o diciembre.

La explicación avanza de lo general a lo particular. Para cada archivo se presenta primero qué informa y qué significa cada indicador; después se documentan la fórmula usada por la macro Excel, el archivo/hoja/celda de origen, la unidad y, cuando el dato fue migrado a Teradata, el query vigente.

**Fórmula Excel:** operación o referencia usada por la macro y sus libros auxiliares.

**Query:** SQL que trae datos directamente de Teradata. El identificador enlaza al SQL correspondiente en la sección [Queries](#5-queries), donde las fechas se muestran con el ejemplo de corte del 30 de junio de 2025.

### Convenciones

| Convención | Significado |
|---|---|
| `TRM` | Tasa Representativa del Mercado. La aplicación la consulta una sola vez por período al servicio web de la Superfinanciera y reutiliza el valor en todos los archivos. Si el servicio falla o entrega un valor inválido, usa `PIB_PEA_TRM_DG` como contingencia. |
| `safeDivide(a,b)` | División que devuelve cero cuando el denominador es nulo o cero. |
| COP / USD | Pesos colombianos / dólares estadounidenses. |
| MM | Millones. |
| `—` | No aplica query Teradata para ese dato. |

## 2. Archivo mensual

El archivo mensual es una serie de tiempo del sistema. Cada fila identifica un mes e informa cobertura, movilidad de afiliados, tamaño y composición del portafolio, rentabilidad, concentración y tipo de cambio.

### 2.1 Qué informa cada columna

| Columna | Concepto | Qué representa / cómo se interpreta |
|---|---|---|
| Afiliados (B) | Afiliados totales | Personas vinculadas al sistema; mide su cobertura acumulada. |
| Aportantes (C) | Aportantes | Afiliados que cotizan; mide cobertura contributiva efectiva. |
| Traspasos anuales (D) | Traspasos | Movimientos entre AFP durante los últimos doce meses. |
| Fondos administrados (E) | Fondos en USD | Tamaño del portafolio administrado convertido a dólares. |
| Inversión total (F) | Portafolio de referencia | Valor total de las inversiones convertido a dólares. |
| Inversión deuda gubernamental (G) | Deuda pública local | Participación del portafolio en deuda gubernamental colombiana. |
| Instrumentos de instituciones financieras (H) | Sector financiero local | Participación en depósitos, efectivo y deuda de instituciones financieras. |
| Instrumentos de instituciones no financieras (I) | Sector no financiero local | Participación en deuda de emisores no financieros. |
| Inversión en acciones (J) | Acciones locales | Exposición a renta variable colombiana. |
| Inversión en fondo mutuo (K) | Fondos locales | Exposición a fondos y vehículos colectivos locales. |
| Instrumentos de emisores extranjeros (L) | Inversión exterior | Suma de las categorías seleccionadas de inversión en el exterior. |
| Otros instrumentos (M) | Otros activos | Categorías del portafolio no incluidas en los rubros anteriores. |
| Rentabilidad nominal (N) | Rentabilidad nominal a un año | Rendimiento de los últimos doce meses sin descontar inflación. |
| Rentabilidad real (O) | Rentabilidad real a un año | Rendimiento de los últimos doce meses después de descontar inflación. |
| Concentración N.° administradoras (P) | Número de administradoras | Número de administradoras vigentes. |
| Concentración fondos administrados (Q) | Concentración de afiliados | Participación de las dos AFP con más afiliados respecto del total. |
| Concentración cuentas administradas (R) | Concentración de fondos | Participación conjunta de Protección y Porvenir en los fondos administrados. |
| Tipo de cambio (S) | TRM | Pesos colombianos por dólar usados en las conversiones del período. |

### 2.2 Matriz de trazabilidad mensual

| Columna | Fórmula Excel | Fuente, hoja, celda | Unidad | Query | Descripción query |
|---|---|---|---|---|---|
| Afiliados (B) | `afiliados = hombres + mujeres` | `Serie_Formato_ 491 AFILIADOS AFP.xlsm`, hoja `informe de prensa`: `C11 + D11`. | Personas | [Q491-TOTALES](#q491-totales) | Suma `TOTAL_AFILIADOS_TOTAL`, renglón 999, para fondos 1000/5000/6000/7000/8000. |
| Aportantes (C) | `aportantes = multifondos!E25` | `Serie_Formato_ 491 AFILIADOS AFP.xlsm`, hoja `multifondos`, celda `E25`. | Personas | [Q491-TOTALES](#q491-totales) | Suma `TOTAL_AFILIADOS_COTIZANTES`, renglón 999, sin filtro por AFP para obtener el sistema. |
| Traspasos anuales (D) | `traspasos_sistema = 'Traslados Entre AFP'!BQ11`, con entidad `99` en `D4` | `Serie_Formato_493 MOVIMIENTO AFILIADOS.xlsx`, hoja `Traslados Entre AFP`: fecha en `B11`, entidad en `D4`, resultado en `BQ11`. | Personas | [Q493-TRASPASOS](#q493-traspasos) | Suma los rangos de edad de hombres y mujeres para las UC/renglones de traspasos en una ventana móvil de doce meses. |
| Fondos administrados (E) | `vr_fondo / TRM` | Macro: balance `SISTEMA TOTAL`, hoja `restot`, total del sistema. Aplicación: saldo PUC `100000` de `ESTFIN_INDIV_PA`. | MM USD | [Q136-FONDOS](#q136-fondos) | Suma el saldo contable de fondos administrados y lo convierte con la TRM. |
| Inversión total (F) | `LIMITES!AIOS!AB4 / TRM` | `LIMITES del nuevo.xlsm`, hoja `AIOS`, celda `AB4`. | USD según escala del libro | — | — |
| Inversión deuda gubernamental (G) | `AIOS!C4 * 100` | `LIMITES del nuevo.xlsm`, hoja `AIOS`, celda `C4`. | Porcentaje | — | — |
| Instituciones financieras (H) | `AIOS!E4 * 100` | `LIMITES del nuevo.xlsm`, hoja `AIOS`, celda `E4`. | Porcentaje | — | — |
| Instituciones no financieras (I) | `AIOS!G4 * 100` | `LIMITES del nuevo.xlsm`, hoja `AIOS`, celda `G4`. | Porcentaje | — | — |
| Acciones (J) | `AIOS!I4 * 100` | `LIMITES del nuevo.xlsm`, hoja `AIOS`, celda `I4`. | Porcentaje | — | — |
| Fondo mutuo (K) | `AIOS!K4 * 100` | `LIMITES del nuevo.xlsm`, hoja `AIOS`, celda `K4`. | Porcentaje | — | — |
| Emisores extranjeros (L) | `(O4 + Q4 + S4 + U4 + W4 + Y4) * 100` | `LIMITES del nuevo.xlsm`, hoja `AIOS`, celdas `O4,Q4,S4,U4,W4,Y4`. | Porcentaje | — | — |
| Otros instrumentos (M) | `AIOS!AA4 * 100` | `LIMITES del nuevo.xlsm`, hoja `AIOS`, celda `AA4`. | Porcentaje | — | — |
| Rentabilidad nominal (N) | `tmp_nominal_1 * 100` | `Rent_Vr_Uni_Moderado.xlsm`: valor nominal de doce meses; la macro carga la serie del período. | Porcentaje | — | — |
| Rentabilidad real (O) | `tmp_real_1 * 100` | `Rent_Vr_Uni_Moderado.xlsm`: valor real de doce meses; equivale a la búsqueda por fecha documentada en la serie. | Porcentaje | — | — |
| N.° administradoras (P) | `4` | Constante de la macro: cuatro administradoras vigentes incluidas en la generación (Colfondos, Porvenir, Protección y Skandia). | Número | — | — |
| Concentración fondos administrados (Q) | `((multifondos!J8 + multifondos!J9) / multifondos!J12) * 100` | Macro: `Serie_Formato_ 491 AFILIADOS AFP.xlsm`, hoja `multifondos`, celdas `J8,J9,J12`. | Porcentaje | [Q491-CONCENTRACION](#q491-concentracion) | Agrupa afiliados por AFP, ordena de mayor a menor y divide las dos mayores entre el total del sistema. |
| Concentración cuentas administradas (R) | `(fondos Protección + fondos Porvenir) / fondos del sistema * 100` | Macro: `SISTEMA TOTAL`, hoja `restot`, valores de Protección, Porvenir y total. Aplicación: PUC `100000` en `ESTFIN_INDIV_PA`. | Porcentaje | [Q136-FONDOS](#q136-fondos) | Calcula los saldos por AFP y obtiene la participación conjunta de Protección y Porvenir. |
| Tipo de cambio (S) | Búsqueda de la TRM aplicable a la fecha | Servicio web TRM de la Superfinanciera; contingencia: `PIB_PEA_TRM_DG`, fecha y valor de TRM. | COP/USD | — | No es Teradata: se consulta una vez al servicio oficial y se reutiliza. |

## 3. Archivo trimestral

El archivo trimestral contiene una hoja por tema. Cada fila corresponde a un cierre de marzo, junio, septiembre (`sep`) o diciembre y se ordena cronológicamente. Las columnas muestran AFP y, cuando aplica, tipo de fondo.

### 3.1 Qué informa cada hoja

| Hoja | Información | Qué representa / cómo se interpreta |
|---|---|---|
| `afiliados` | Número de afiliados por administradora y fondo | Distribución de las personas entre AFP, fondos moderado, conservador, mayor riesgo y convergencias. |
| `aportantes` | Número de aportantes por administradora | Base contributiva activa de cada AFP. |
| `colombia` | Fondos administrados por AFP y fondo | Saldos al cierre expresados en millones de dólares. |
| `gastos` | Gastos operativos por administradora | Administración, comercialización y otros gastos de los últimos doce meses. |
| `comisiones` | Estructura de comisiones | Comisión sobre flujo y seguro vigentes al cierre. |
| `rentabilidad` | Rendimiento nominal y real | Desempeño anual de los últimos doce meses por AFP para el fondo moderado. |
| `promotores` | Número de promotores | Campo histórico; se informa `n.d.` cuando no existe fuente disponible. |
| `traspasos` | Traspasos entre administradoras | Movilidad acumulada durante los últimos doce meses por AFP. |

### 3.2 Matriz de trazabilidad trimestral

| Columna / hoja | Fórmula Excel | Fuente, hoja, celda | Unidad | Query | Descripción query |
|---|---|---|---|---|---|
| Afiliados por AFP y fondo | La macro toma las celdas `C:H` de las filas 8–11 de `multifondos`; para Skandia moderado suma `mod_sk + alt_sk`. | `Serie_Formato_ 491 AFILIADOS AFP.xlsm`, hoja `multifondos`: Porvenir fila 8, Protección 9, Colfondos 10, Skandia 11. | Personas | [Q491-AFILIADOS-FONDO](#q491-afiliados-fondo) | Agrupa `TOTAL_AFILIADOS_TOTAL` por AFP, UC y código de fondo con renglón 999. |
| Aportantes por AFP | `cot_colf`, `cot_porv`, `cot_prot`, `cot_sk`; las administradoras históricas sin dato reciben cero. | `Serie_Formato_ 491 AFILIADOS AFP.xlsm`, `multifondos!J19:J22`. | Personas | [Q491-APORTANTES-ENTIDAD](#q491-aportantes-entidad) | Suma `TOTAL_AFILIADOS_COTIZANTES` para cada código de AFP. |
| Fondos administrados | `saldo del fondo y AFP / TRM`; Skandia moderado suma el fondo alternativo. | Macro: balances por tipo de fondo (`MOD`, `CON`, `MR`, `RP`) y `SISTEMA TOTAL`. | MM USD | [Q136-FONDOS](#q136-fondos) | Suma PUC `100000` por patrimonio y AFP en `ESTFIN_INDIV_PA`, luego convierte con TRM. |
| Gastos operativos | `(débito - crédito) / TRM` para cada AFP. | `Plantilla AIOS-probable.xlsm`, hoja `cuentas`: Protección `C50-D57`, Porvenir `C51-D69`, Skandia `C52-D81`, Colfondos `C53-D93`; datos originados en `base anual`. | USD según escala del boletín | — | — |
| Comisiones | Cada comisión obligatoria y de seguro se multiplica por `100`. | `Comisión FPO desde 2003.xlsx`, hoja `COTIZACION CORTE ANUAL`: Skandia `B1:C1`, Porvenir `F1:G1`, Protección `N1:O1`, Colfondos `R1:S1`. | Porcentaje | — | — |
| Rentabilidad nominal/real | `tmp_nominal_AFP_12 * 100` y `tmp_real_AFP_12 * 100`. | Libros de rentabilidad por AFP y fondo; series del cierre y de doce meses antes. | Porcentaje | — | — |
| Promotores | `n.d.` | Sin fuente implementada; se conserva explícitamente como no disponible. | No aplica | — | — |
| Traspasos por AFP | Resultado de `BQ11` después de escribir la fecha en `B11` y el código de AFP en `D4`. | `Serie_Formato_493 MOVIMIENTO AFILIADOS.xlsx`, hoja `Traslados Entre AFP`. | Personas | [Q493-TRASPASOS](#q493-traspasos) | Misma ventana y suma de rangos del sistema, agregando `CODIGO_ENTIDAD` por AFP. |

## 4. Archivo semestral

El archivo semestral reúne indicadores de cobertura, demografía, pensionados, movilidad, tamaño financiero, composición del portafolio, situación contable, eficiencia, comisiones y rentabilidad. Cada columna representa un cierre de junio o diciembre.

| Filas | Bloque informado |
|---:|---|
| 3–15 | Afiliados, edades, aportantes, cobertura frente a la PEA y salario promedio. |
| 16–24 | Pensionados y altas de beneficiarios por tipo de prestación. |
| 25–27 | Fallecimientos y traspasos. |
| 28–47 | Fondos administrados, tamaño relativo y composición de inversiones. |
| 48–69 | Balances, resultados, gastos e indicadores de eficiencia. |
| 70–80 | Aportes, comisiones y antigüedad del sistema. |
| 82–89 | Rentabilidades nominales y reales de 10, 5, 3 y 1 año. |

### 4.1 Diccionario conceptual de datos semestrales

| Fila | Indicador | Qué representa / cómo se interpreta |
|---:|---|---|
| 3 | Afiliados activos | Personas afiliadas activas; mide la población con vínculo vigente al sistema. |
| 4–7 | Distribución por edad | Participación de afiliados menores de 30, de 30–44, de 45–59 y mayores de 60 años. |
| 8 | Total por edad | Control de que la distribución etaria suma 100 %. |
| 9 | Afiliados en miles | Tamaño del sistema expresado en una escala comparable internacionalmente. |
| 10 | Participación de mujeres | Proporción de mujeres dentro del total de afiliados. |
| 11 | Aportantes | Personas que cotizan efectivamente. |
| 12–14 | Cobertura y densidad contributiva | Afiliados/PEA, aportantes/PEA y aportantes/afiliados. |
| 15 | Salario promedio en USD | Ingreso ponderado de los afiliados convertido a dólares. |
| 16–19 | Pensionados | Total y distribución por invalidez, vejez y sobrevivencia. |
| 20–24 | Altas de beneficiarios | Bloque previsto por la plantilla; la aplicación actual no lo genera y conserva las celdas de referencia. |
| 25 | Fallecimiento anual | Afiliados fallecidos durante la ventana anual, expresados en miles. |
| 26–27 | Traspasos | Total anual y proporción respecto de afiliados. |
| 28–29 | Fondos administrados | Tamaño en MM USD y relación frente al PIB. |
| 30–44 | Composición del portafolio | Participaciones locales, exteriores, titulizadoras, otros e inversión en moneda extranjera. |
| 45 | Fondos / deuda gubernamental | Tamaño de los fondos frente al saldo de deuda pública. |
| 46–47 | Administración y concentración | Número de AFP y participación de las dos mayores. |
| 48–50 | Balance | Activo, pasivo y patrimonio neto en dólares. |
| 51–54 | Resultados | Comisiones, gastos, resultado operacional y resultado neto. |
| 55–60 | Gastos | Administración, comercialización y otros gastos. |
| 61–69 | Eficiencia | Recaudación por aportante y razones de gastos, patrimonio, utilidad y comisiones. |
| 70–76 | Aportes y comisión | Tasa obligatoria y distribución entre trabajador, empleador y Estado. |
| 77–80 | Comisiones/fondos y antigüedad | Peso de comisiones sobre fondos y años transcurridos desde 1994. |
| 82–89 | Rentabilidad | Retornos nominales y reales anualizados para 10, 5, 3 y 1 año. |

### 4.2 Matriz de trazabilidad semestral: población, pensionados y movilidad

| Fila / indicador | Fórmula Excel | Fuente, hoja, celda | Unidad | Query | Descripción query |
|---|---|---|---|---|---|
| 3. Afiliados activos | Macro histórica: `afiliados = hombres + mujeres`; la salida vigente usa afiliados activos. | Macro: `Serie_Formato_ 491 AFILIADOS AFP.xlsm`, `informe de prensa!C11:D11`. | Personas | [Q491-TOTALES](#q491-totales) | Suma `TOTAL_AFILIADOS_ACTIVOS_TOTAL`, renglón 999 y fondos definidos. |
| 4. Afiliados <30 | `(informe de prensa!C81 + D81) / afiliados * 100` | `Serie_Formato_ 491 AFILIADOS AFP.xlsm`, `informe de prensa!C81:D81`. | Porcentaje | [Q491-EDADES](#q491-edades) | Suma afiliados de UC 1 con renglón menor que 80. |
| 5. Afiliados 30–44 | `(C82 + D82) / afiliados * 100` | Mismo libro, `informe de prensa!C82:D82`. | Porcentaje | [Q491-EDADES](#q491-edades) | Aplica reglas de UC/renglón del grupo 30–44. |
| 6. Afiliados 45–59 | `(C83 + D83) / afiliados * 100` | Mismo libro, `informe de prensa!C83:D83`. | Porcentaje | [Q491-EDADES](#q491-edades) | Aplica reglas de UC/renglón del grupo 45–59. |
| 7. Afiliados >60 | `(C84 + D84) / afiliados * 100` | Mismo libro, `informe de prensa!C84:D84`. | Porcentaje | [Q491-EDADES](#q491-edades) | Aplica reglas de UC/renglón del grupo mayor de 60. |
| 8. Total | Suma de los porcentajes de filas 4–7; la aplicación escribe `100`. | Filas 4–7 del semestral. | Porcentaje | — | — |
| 9. Afiliados miles | `afiliados / 1000` | Total de afiliados. | Miles de personas | [Q491-TOTALES](#q491-totales) | Reutiliza el total de afiliados consultado. |
| 10. Participación mujeres | `mujeres / afiliados * 100` | Macro: `informe de prensa!D11 / (C11+D11)`. | Porcentaje | [Q491-TOTALES](#q491-totales) | Suma `TOTAL_AFILIADOS_M` y divide entre afiliados. |
| 11. Aportantes | `multifondos!E25` | `Serie_Formato_ 491 AFILIADOS AFP.xlsm`, `multifondos!E25`. | Personas | [Q491-TOTALES](#q491-totales) | Suma `TOTAL_AFILIADOS_COTIZANTES` del sistema. |
| 12. Afiliados / PEA | `afiliados / PEA * 100` | Afiliados y `PIB_PEA_TRM_DG`, serie PEA aplicable al corte. | Porcentaje | [Q491-TOTALES](#q491-totales) | El numerador proviene del Formato 491; PEA continúa en archivo. |
| 13. Aportantes / PEA | `aportantes / PEA * 100` | Aportantes y `PIB_PEA_TRM_DG`, serie PEA. | Porcentaje | [Q491-TOTALES](#q491-totales) | El numerador proviene del Formato 491. |
| 14. Aportantes / afiliados | `aportantes / afiliados * 100` | Totales de aportantes y afiliados. | Porcentaje | [Q491-TOTALES](#q491-totales) | Reutiliza ambos agregados del Formato 491. |
| 15. Salario promedio | `SM COLOMBIA!E8 / TRM` | Macro: `Serie_Formato_ 491 AFILIADOS AFP.xlsm`, hoja `SM COLOMBIA`, celda `E8`. | USD | [Q491-SALARIO](#q491-salario) | Pondera rangos IBC por salario mínimo y divide por afiliados; luego convierte con TRM. |
| 16. Total pensionados | `por entidad!BJ67` | `Series_Formato-495 PENSIONADOS.xlsm`, hoja `por entidad`, celda `BJ67`. | Personas | [Q495-PENSIONADOS](#q495-pensionados) | Suma las columnas de pensiones con UC 1 y renglón 200. |
| 17. Invalidez | `por entidad!BI66 / fila 16` | Mismo libro, `por entidad!BI66`. | Porcentaje | [Q495-PENSIONADOS](#q495-pensionados) | Suma las columnas de invalidez y divide por total. |
| 18. Vejez | `por entidad!BH66 / fila 16` | Mismo libro, `por entidad!BH66`. | Porcentaje | [Q495-PENSIONADOS](#q495-pensionados) | Suma las columnas de vejez y divide por total. |
| 19. Sobrevivencia | `por entidad!BJ66 / fila 16` | Mismo libro, `por entidad!BJ66`. | Porcentaje | [Q495-PENSIONADOS](#q495-pensionados) | Suma las columnas de sobrevivencia y divide por total. |
| 20–24. Altas de beneficiarios | La macro escribe `"no disponible"` en la fila 20; no calcula las filas 21–24. | Sin fuente implementada. | No disponible | — | — |
| 25. Fallecimientos | `'Fallecidos'!M11 / 1000`, con fecha en `B11` y entidad 99 en `D4`. | `Serie_Formato_493 MOVIMIENTO AFILIADOS.xlsx`, hoja `Fallecidos`. | Miles de personas | [Q493-FALLECIDOS](#q493-fallecidos) | Suma rangos de sexo/edad, UC 1, renglones 165/170/175 y ventana anual. |
| 26. Traspasos anuales | `'Traslados Entre AFP'!BQ11` con entidad 99. | `Serie_Formato_493 MOVIMIENTO AFILIADOS.xlsx`, hoja `Traslados Entre AFP`. | Personas | [Q493-TRASPASOS](#q493-traspasos) | Total de traspasos del sistema en la ventana anual. |
| 27. Traspasos / afiliados | `traspasos_sistema / afiliados` | Filas 26 y total de afiliados. | Porcentaje mediante formato Excel | [Q493-TRASPASOS](#q493-traspasos) | Reutiliza traspasos; el denominador viene de Q491-TOTALES. |

### 4.3 Matriz de trazabilidad semestral: fondos y composición

| Fila / indicador | Fórmula Excel | Fuente, hoja, celda | Unidad | Query | Descripción query |
|---|---|---|---|---|---|
| 28. Fondos administrados | `vr_fondo / TRM` | Macro: `SISTEMA TOTAL`, `restot`, total sistema. | MM USD | [Q136-FONDOS](#q136-fondos) | Suma saldos PUC `100000` y convierte con TRM. |
| 29. Fondos / PIB | `(vr_fondo / TRM) / (PIB / TRM)` | `SISTEMA TOTAL` y serie PIB de `PIB_PEA_TRM_DG`. | Porcentaje mediante formato Excel | [Q136-FONDOS](#q136-fondos) | El valor de fondos procede de Teradata; PIB permanece en archivo. |
| 30. Composición total | `LIMITES!AIOS!AB4 / TRM` | `LIMITES del nuevo.xlsm`, `AIOS!AB4`. | USD según escala | — | — |
| 31. Deuda gubernamental local | `AIOS!C4` | `LIMITES del nuevo.xlsm`, `AIOS!C4`. | Ratio/porcentaje | — | — |
| 32. Instituciones financieras locales | `AIOS!E4` | Mismo libro, `AIOS!E4`. | Ratio/porcentaje | — | — |
| 33. Instituciones no financieras locales | `AIOS!G4` | Mismo libro, `AIOS!G4`. | Ratio/porcentaje | — | — |
| 34. Acciones locales | `AIOS!I4` | Mismo libro, `AIOS!I4`. | Ratio/porcentaje | — | — |
| 35. Administradores de fondos locales | `AIOS!K4` | Mismo libro, `AIOS!K4`. | Ratio/porcentaje | — | — |
| 36. Sociedades titulizadoras locales | `0` | Constante vigente. | Ratio/porcentaje | — | — |
| 37. Deuda gubernamental exterior | `AIOS!O4` | `LIMITES del nuevo.xlsm`, `AIOS!O4`. | Ratio/porcentaje | — | — |
| 38. Instituciones financieras exterior | `AIOS!Q4` | Mismo libro, `AIOS!Q4`. | Ratio/porcentaje | — | — |
| 39. Instituciones no financieras exterior | `AIOS!S4` | Mismo libro, `AIOS!S4`. | Ratio/porcentaje | — | — |
| 40. Acciones exterior | `AIOS!U4` | Mismo libro, `AIOS!U4`. | Ratio/porcentaje | — | — |
| 41. Administradores de fondos exterior | `AIOS!W4` | Mismo libro, `AIOS!W4`. | Ratio/porcentaje | — | — |
| 42. Sociedades titulizadoras exterior | Macro: `AIOS!Y4`; aplicación vigente: constante `2`. | `LIMITES del nuevo.xlsm`, `AIOS!Y4`; se documenta la diferencia respecto del valor vigente. | Ratio/porcentaje | — | — |
| 43. Otros | `AIOS!AA4` | `LIMITES del nuevo.xlsm`, `AIOS!AA4`. | Ratio/porcentaje | — | — |
| 44. Inversión en moneda extranjera | `O4 + Q4 + S4 + U4 + W4 + Y4` | `LIMITES del nuevo.xlsm`, hoja `AIOS`. | Ratio/porcentaje | — | — |
| 45. Fondos / deuda gubernamental | `(vr_fondo / TRM) / deuda_gubernamental_USD` | Fondos del sistema y `PIB_PEA_TRM_DG`, serie de deuda gubernamental. | Porcentaje mediante formato Excel | [Q136-FONDOS](#q136-fondos) | Fondos desde Teradata; deuda continúa en archivo. |
| 46. Número de administradoras | `4` | Constante: Colfondos, Porvenir, Protección y Skandia. | Número | — | — |
| 47. Participación de las dos mayores | `(fondos Protección + fondos Porvenir) / total fondos * 100` | Macro: `SISTEMA TOTAL`, `restot`; aplicación: saldos PUC `100000`. | Porcentaje | [Q136-FONDOS](#q136-fondos) | Calcula la participación conjunta de Protección y Porvenir. |

### 4.4 Matriz de trazabilidad semestral: balance, gastos y eficiencia

| Fila / indicador | Fórmula Excel | Fuente, hoja, celda | Unidad | Query | Descripción query |
|---|---|---|---|---|---|
| 48. Activo | `CUENTAS!C6 / TRM` | `Plantilla AIOS-probable.xlsm`, hoja `CUENTAS`, celda `C6`. | USD | — | — |
| 49. Pasivo | `CUENTAS!C4 / TRM` | Mismo libro, `CUENTAS!C4`. | USD | — | — |
| 50. Patrimonio neto | `(CUENTAS!C6 - CUENTAS!C4) / TRM` | Mismo libro, activo `C6` y pasivo `C4`. | USD | — | — |
| 51. Ingresos por comisiones | `CUENTAS!E13` | `Plantilla AIOS-probable.xlsm`, `CUENTAS!E13`. | Unidad contable de la plantilla | — | — |
| 52. Gastos operativos | `CUENTAS!G15` | Mismo libro, `CUENTAS!G15`. | Unidad contable de la plantilla | — | — |
| 53. Resultado operativo | Macro: `comisiones - gastos`; aplicación: valor contable de resultado operativo. | Macro usa filas 51–52; aplicación lee `CUENTAS!E41`. | Unidad contable de la plantilla | — | — |
| 54. Resultado neto | `CUENTAS!E44` | `Plantilla AIOS-probable.xlsm`, `CUENTAS!E44`. | Unidad contable de la plantilla | — | — |
| 55. Gastos de administración | `CUENTAS!H24` | Mismo libro, `CUENTAS!H24`. | Unidad contable de la plantilla | — | — |
| 56. Comisión vendedores | `cuenta 511500 / TRM` | `Plantilla AIOS-probable.xlsm`, hoja `cuentas`, celda `C21`. | USD | — | — |
| 57. Comercialización | `cuenta 511527 / TRM` | Mismo libro, `cuentas!C22`. | USD | — | — |
| 58. Total comercialización | `(C21 + C22) / TRM` | Mismo libro, `cuentas!C21:C22`. | USD | — | — |
| 59. Otros gastos | `(C24+C28+C29+C31+C32+C33+C34+C35+C36+C37+C38) / TRM` | `Plantilla AIOS-probable.xlsm`, hoja `cuentas`; cuentas 512000, 513000, 513500, 514000, 514500, 515000, 515500, 516000, 516500, 517000 y 517200. | USD | — | — |
| 60. Total gastos | `cuenta 510000 / TRM` | Mismo libro, `cuentas!C15`. | USD | — | — |
| 61. Recaudación anual por aportante | `(FORMATO OBL!E6 / TRM) / (aportantes/1000) * 1000` | Macro: `Formato_136_Meses.xlsm`, hoja `FORMATO OBL`, fechas en `B6:B7`, resultado `E6`. | USD por mil aportantes | [Q136-APORTES](#q136-aportes) | Suma aportes recibidos del Formato 136 entre 2024-06-01 y 2025-06-30; sustituye la lectura de `E6`. |
| 62. Gastos / recaudación | `gastos / (aportes_recibidos / TRM) * 100` | Gastos de fila 52 y aportes que en la macro provienen de `FORMATO OBL!E6`. | Porcentaje | [Q136-APORTES](#q136-aportes) | Reutiliza los aportes recibidos consultados. |
| 63. Patrimonio / fondos | `(patrimonio base mes / TRM) / fila 28 * 100` | `Plantilla AIOS-probable.xlsm`, `base mes`, patrimonio del período; fila 28. | Porcentaje | — | — |
| 64. Patrimonio por afiliado | `fila 50 / afiliados * 1,000,000` | Patrimonio en USD de fila 50 y afiliados. | USD por afiliado | [Q491-TOTALES](#q491-totales) | El denominador proviene del total de afiliados. |
| 65. Utilidad / comisiones | `resultado_neto / comisiones * 100` | `CUENTAS!E44 / E13`. | Porcentaje | — | — |
| 66. Utilidad / patrimonio | `resultado_neto / patrimonio_USD * 100` | `CUENTAS!E44` y fila 50. | Porcentaje | — | — |
| 67. Gastos por afiliado | `gastos / afiliados * 1,000,000` | `CUENTAS!G15` y afiliados. | Valor por afiliado | [Q491-TOTALES](#q491-totales) | El denominador proviene del total de afiliados. |
| 68. Comisiones por aportante | `comisiones / aportantes * 1,000,000` | `CUENTAS!E13` y aportantes. | Valor por aportante | [Q491-TOTALES](#q491-totales) | El denominador proviene del total de aportantes. |
| 69. Comisión / recaudación neta | `gastos_administración / fila 61` | `CUENTAS!H24` y fila 61. | Ratio | [Q136-APORTES](#q136-aportes) | La recaudación normalizada usa aportes del Formato 136. |

### 4.5 Matriz de trazabilidad semestral: aportes, comisiones y rentabilidad

| Fila / indicador | Fórmula Excel | Fuente, hoja, celda | Unidad | Query | Descripción query |
|---|---|---|---|---|---|
| 70. Tasa de aporte obligatorio | `16` | Constante regulatoria de la macro. | Porcentaje | — | — |
| 71. Comisión sobre el salario | `promedio(comisión obligatoria de Colfondos, Porvenir, Protección y Skandia) * 100` | `Comisión FPO desde 2003.xlsx`, `COTIZACION CORTE ANUAL`: `B1,F1,N1,R1`. | Porcentaje | — | — |
| 72. Comisión sobre saldo | `0` | Constante de la macro. | Porcentaje | — | — |
| 73. Comisión sobre rentabilidad | `0` | Constante de la macro. | Porcentaje | — | — |
| 74. Porcentaje trabajador | `(3 - fila 71) * 0.25` | Fila 71 y constantes 3/0,25. | Porcentaje | — | — |
| 75. Porcentaje empleador | `(3 - fila 71) * 0.75` | Fila 71 y constantes 3/0,75. | Porcentaje | — | — |
| 76. Porcentaje Estado | `0` | Constante de la macro. | Porcentaje | — | — |
| 77. Ingresos por comisiones (a) | `CUENTAS!E13` | `Plantilla AIOS-probable.xlsm`, `CUENTAS!E13`. | Unidad contable | — | — |
| 78. Fondo de aportes obligatorios (b) | Reutiliza `fila 28` | Fondos administrados en MM USD. | MM USD | [Q136-FONDOS](#q136-fondos) | Reutiliza los fondos calculados desde PUC `100000`. |
| 79. (a)/(b) | `fila 77 / fila 78` | Filas 77 y 78. | Ratio | — | — |
| 80. Antigüedad en el sistema | `año(fecha_corte) - 1994` | Fecha del período. | Años | — | — |
| 82. Últimos 10 años nominal | `tmp_nominal_10 * 100` | NAV de `Valores_Fondo_Moder` / hoja `MODERADO`, valores inicial y final. | Porcentaje anualizado | — | — |
| 83. Últimos 10 años real | `tmp_real_10 * 100` | Nominal de 10 años e IPC de `Rent_Vr_Uni_Moderado`. | Porcentaje anualizado | — | — |
| 84. Últimos 5 años nominal | `tmp_nominal_5 * 100` | NAV de `Valores_Fondo_Moder` / `MODERADO`. | Porcentaje anualizado | — | — |
| 85. Últimos 5 años real | `tmp_real_5 * 100` | Nominal de 5 años e IPC. | Porcentaje anualizado | — | — |
| 86. Últimos 3 años nominal | `tmp_nominal_3 * 100` | NAV de `Valores_Fondo_Moder` / `MODERADO`. | Porcentaje anualizado | — | — |
| 87. Últimos 3 años real | `tmp_real_3 * 100` | Nominal de 3 años e IPC. | Porcentaje anualizado | — | — |
| 88. Últimos 12 meses nominal | `tmp_nominal_1 * 100` | NAV de `Valores_Fondo_Moder` / `MODERADO`. | Porcentaje | — | — |
| 89. Últimos 12 meses real | `tmp_real_1 * 100` | Nominal de 12 meses e IPC de `Rent_Vr_Uni_Moderado`. | Porcentaje | — | — |

## 5. Queries

Los SQL siguientes usan fechas literales para que el ejemplo del **30 de junio de 2025** sea legible. La aplicación ejecuta las mismas sentencias con parámetros preparados (`?`). Los enlaces al código Java permiten auditar la versión exacta que se ejecuta.

<a id="q491-totales"></a>
### Q491-TOTALES — afiliados, activos, mujeres y aportantes

La aplicación ejecuta un agregado por métrica. Se muestran juntos porque comparten fecha, renglón y fondos:

```sql
SELECT
  COALESCE(SUM(COALESCE(TOTAL_AFILIADOS_TOTAL,0)),0) AS afiliados,
  COALESCE(SUM(COALESCE(TOTAL_AFILIADOS_ACTIVOS_TOTAL,0)),0) AS afiliados_activos,
  COALESCE(SUM(COALESCE(TOTAL_AFILIADOS_M,0)),0) AS mujeres,
  COALESCE(SUM(COALESCE(TOTAL_AFILIADOS_COTIZANTES,0)),0) AS aportantes
FROM PROD_DWH_CONSULTA.FORMATO491
WHERE FECBAL = DATE '2025-06-30'
  AND RENGLON = '999'
  AND SUBSTR(NUMERO_IDENTIFICACION,9,4)
      IN ('1000','5000','6000','7000','8000');
```

Código ejecutable: [`Formato491QueryService`](../src/main/java/co/gov/sfc/excel/Formato491QueryService.java).

<a id="q491-edades"></a>
### Q491-EDADES — rangos de edad

```sql
SELECT
 SUM(CASE WHEN CAST(TRIM(UNIDAD_CAPTURA) AS INTEGER)=1
                AND CAST(TRIM(RENGLON) AS INTEGER)<80
          THEN COALESCE(TOTAL_AFILIADOS_TOTAL,0) ELSE 0 END) AS menor_30,
 SUM(CASE WHEN (CAST(TRIM(UNIDAD_CAPTURA) AS INTEGER)=1
                AND CAST(TRIM(RENGLON) AS INTEGER) BETWEEN 80 AND 150)
               OR (CAST(TRIM(UNIDAD_CAPTURA) AS INTEGER)=4
                AND CAST(TRIM(RENGLON) AS INTEGER) BETWEEN 5 AND 15)
          THEN COALESCE(TOTAL_AFILIADOS_TOTAL,0) ELSE 0 END) AS edad_30_44,
 SUM(CASE WHEN CAST(TRIM(RENGLON) AS INTEGER) BETWEEN 155 AND 225
               OR (CAST(TRIM(UNIDAD_CAPTURA) AS INTEGER)>1
                AND CAST(TRIM(RENGLON) AS INTEGER) BETWEEN 20 AND 50)
               OR (CAST(TRIM(UNIDAD_CAPTURA) AS INTEGER) BETWEEN 2 AND 3
                AND CAST(TRIM(RENGLON) AS INTEGER)<20)
          THEN COALESCE(TOTAL_AFILIADOS_TOTAL,0) ELSE 0 END) AS edad_45_59,
 SUM(CASE WHEN (CAST(TRIM(RENGLON) AS INTEGER)>=230
                AND CAST(TRIM(RENGLON) AS INTEGER)<999)
               OR (CAST(TRIM(UNIDAD_CAPTURA) AS INTEGER)>1
                AND CAST(TRIM(RENGLON) AS INTEGER) BETWEEN 55 AND 80)
          THEN COALESCE(TOTAL_AFILIADOS_TOTAL,0) ELSE 0 END) AS mayor_60
FROM PROD_DWH_CONSULTA.FORMATO491
WHERE FECBAL = DATE '2025-06-30'
  AND SUBSTR(NUMERO_IDENTIFICACION,9,4)
      IN ('1000','5000','6000','7000','8000');
```

La aplicación conserva cuatro consultas separadas para facilitar la trazabilidad de cada grupo.

<a id="q491-afiliados-fondo"></a>
### Q491-AFILIADOS-FONDO — afiliados trimestrales por AFP y fondo

```sql
SELECT CODIGO_ENTIDAD,
       CAST(TRIM(UNIDAD_CAPTURA) AS INTEGER) AS unidad_captura,
       SUBSTR(NUMERO_IDENTIFICACION,9,4) AS fondo,
       COALESCE(SUM(COALESCE(TOTAL_AFILIADOS_TOTAL,0)),0) AS afiliados
FROM PROD_DWH_CONSULTA.FORMATO491
WHERE FECBAL = DATE '2025-06-30'
  AND RENGLON = '999'
  AND CODIGO_ENTIDAD IN (2,3,9,10)
  AND (
       (CAST(TRIM(UNIDAD_CAPTURA) AS INTEGER)=1
        AND SUBSTR(NUMERO_IDENTIFICACION,9,4) IN ('1000','5000','6000','8000'))
    OR (CAST(TRIM(UNIDAD_CAPTURA) AS INTEGER) IN (2,3)
        AND SUBSTR(NUMERO_IDENTIFICACION,9,4)='5000')
    OR (CAST(TRIM(UNIDAD_CAPTURA) AS INTEGER)=4
        AND SUBSTR(NUMERO_IDENTIFICACION,9,4)='1000')
  )
GROUP BY 1,2,3;
```

<a id="q491-aportantes-entidad"></a>
### Q491-APORTANTES-ENTIDAD — aportantes trimestrales por AFP

```sql
SELECT CODIGO_ENTIDAD,
       COALESCE(SUM(COALESCE(TOTAL_AFILIADOS_COTIZANTES,0)),0) AS aportantes
FROM PROD_DWH_CONSULTA.FORMATO491
WHERE FECBAL = DATE '2025-06-30'
  AND RENGLON = '999'
  AND CODIGO_ENTIDAD IN (2,3,9,10)
  AND SUBSTR(NUMERO_IDENTIFICACION,9,4)
      IN ('1000','5000','6000','7000','8000')
GROUP BY CODIGO_ENTIDAD;
```

Los códigos son Protección `2`, Porvenir `3`, Skandia `9` y Colfondos `10`.

<a id="q491-concentracion"></a>
### Q491-CONCENTRACION — concentración de afiliados

```sql
SELECT CODIGO_ENTIDAD,
       COALESCE(SUM(COALESCE(TOTAL_AFILIADOS_TOTAL,0)),0) AS afiliados
FROM PROD_DWH_CONSULTA.FORMATO491
WHERE FECBAL = DATE '2025-06-30'
  AND RENGLON = '999'
  AND SUBSTR(NUMERO_IDENTIFICACION,9,4)
      IN ('1000','5000','6000','7000','8000')
GROUP BY CODIGO_ENTIDAD;
```

Java ordena el resultado de mayor a menor y calcula `(AFP1 + AFP2) / total_sistema * 100`.

<a id="q491-salario"></a>
### Q491-SALARIO — salario promedio ponderado

El query pondera los rangos IBC de hombres y mujeres por 1, 2, 3, 4, 8, 12, 16, 20 y 25 salarios mínimos, aplica las combinaciones de UC/fondo del Formato 491 y divide por afiliados. El salario oficial 2025 procede de `SalarioMinimo.csv`.

Código y SQL completo: [`Formato491QueryService.sqlSalarioMinimoPonderado`](../src/main/java/co/gov/sfc/excel/Formato491QueryService.java).

<a id="q493-traspasos"></a>
### Q493-TRASPASOS — traspasos del sistema o por AFP

Para junio de 2025, la ventana anual usada por la aplicación es del 31 de julio de 2024 al 30 de junio de 2025.

```sql
SELECT COALESCE(SUM(
 COALESCE(MUJERES_RANGO_EDAD_31,0)+COALESCE(MUJERES_RANGO_EDAD_31_36,0)+
 COALESCE(MUJERES_RANGO_EDAD_36_41,0)+COALESCE(MUJERES_RANGO_EDAD_41_46,0)+
 COALESCE(MUJERES_RANGO_EDAD_46,0)+COALESCE(HOMBRES_RANGO_EDAD_36,0)+
 COALESCE(HOMBRES_RANGO_EDAD_36_41,0)+COALESCE(HOMBRES_RANGO_EDAD_41_46,0)+
 COALESCE(HOMBRES_RANGO_EDAD_46_51,0)+COALESCE(HOMBRES_RANGO_EDAD_51,0)),0)
  AS total_personas
FROM PROD_DWH_CONSULTA.S9_FORMATO_493
WHERE FECHA_CORTE BETWEEN DATE '2024-07-31' AND DATE '2025-06-30'
  AND ((UNIDAD_CAPTURA=1 AND RENGLON IN (70,75,90,95))
    OR (UNIDAD_CAPTURA=2 AND RENGLON IN (40,45,60,65))
    OR (UNIDAD_CAPTURA=3 AND RENGLON IN (40,45,60,65))
    OR (UNIDAD_CAPTURA=6 AND RENGLON IN (35,40,45,50)));
```

Para una AFP se agrega `CODIGO_ENTIDAD = 2|3|9|10`. Código: [`Formato493QueryService`](../src/main/java/co/gov/sfc/excel/Formato493QueryService.java).

<a id="q493-fallecidos"></a>
### Q493-FALLECIDOS — fallecimientos del sistema

Usa la misma suma de columnas sexo/edad de Q493-TRASPASOS con este filtro:

```sql
FROM PROD_DWH_CONSULTA.S9_FORMATO_493
WHERE UNIDAD_CAPTURA = 1
  AND RENGLON IN (165,170,175)
  AND FECHA_CORTE BETWEEN DATE '2024-07-31' AND DATE '2025-06-30';
```

<a id="q136-fondos"></a>
### Q136-FONDOS — fondos administrados por AFP y patrimonio

```sql
SELECT e.Codigo_Entidad,
       SUM(eip.Saldo_Sincierre_Total_Moneda_0)/1000 AS valor_miles
FROM PROD_DWH_CONSULTA.ESTFIN_INDIV_PA eip
JOIN PROD_DWH_CONSULTA.ENTIDADES e ON eip.Ent_ID=e.Ent_ID
JOIN PROD_DWH_CONSULTA.PATRIMONIOS_AUTONOMOS pa ON eip.Paau_ID=pa.Paau_ID
JOIN PROD_DWH_CONSULTA.TIEMPO t ON eip.Tie_ID=t.Tie_ID
JOIN PROD_DWH_CONSULTA.PUC p ON eip.Puc_ID=p.Puc_ID
WHERE eip.Tipo_Informe=17 AND e.Tipo_Entidad=23 AND e.Estado=1
  AND pa.Tipo_Patrimonio=6 AND pa.Codigo_Patrimonio=1000
  AND p.Codigo=100000 AND t.Fecha=DATE '2025-06-30'
GROUP BY 1;
```

Para los demás fondos se sustituye el código de patrimonio; Skandia moderado incluye patrimonios `4` y `8000`. Código: [`Formato136QueryService`](../src/main/java/co/gov/sfc/excel/Formato136QueryService.java) y [`FondoAdministradoQueryService`](../src/main/java/co/gov/sfc/excel/FondoAdministradoQueryService.java).

<a id="q136-aportes"></a>
### Q136-APORTES — aportes recibidos

La ventana para el corte de ejemplo va del 1 de junio de 2024 al 30 de junio de 2025.

```sql
SELECT COALESCE(SUM(e.valor)/1000000,0) AS valor_total
FROM prod_dwh_consulta.entidades a,
     prod_dwh_consulta.tiempo b,
     prod_dwh_consulta.patrimonios_autonomos c,
     prod_dwh_consulta.negfid_insumos d,
     prod_dwh_consulta.negfid_insumo_entidad e
WHERE d.inf_id=e.inf_id AND e.ent_id=a.ent_id AND e.tie_id=b.tie_id
  AND e.paau_id=c.paau_id
  AND c.tipo_patrimonio=6 AND c.codigo_patrimonio=1000
  AND d.nivel1=136 AND d.nivel2=2 AND d.nivel3=4 AND d.nivel4=10
  AND a.tipo_entidad=23 AND e.valor<>0
  AND b.fecha BETWEEN DATE '2024-06-01' AND DATE '2025-06-30';
```

Este query reemplaza la lectura de `Formato_136_Meses.xlsm`, hoja `FORMATO OBL`, celda `E6`.

<a id="q495-pensionados"></a>
### Q495-PENSIONADOS — total y composición de pensionados

Las consultas de total, invalidez, vejez y sobrevivencia suman las columnas específicas de cada prestación y comparten el siguiente filtro:

```sql
SELECT COALESCE(
  /* suma de las columnas de la prestación solicitada */
,0) AS total_indicador
FROM PROD_DWH_CONSULTA.S9_FORMATO_495
WHERE FECHA_CORTE = DATE '2025-06-30'
  AND UNIDAD_CAPTURA = 1
  AND RENGLON = 200;
```

La lista exacta de columnas de cada variante está en [`Formato495QueryService`](../src/main/java/co/gov/sfc/excel/Formato495QueryService.java), métodos `sqlTotal`, `sqlInvalidez`, `sqlVejez` y `sqlSobrevivencia`. Esa clase es la fuente de verdad para evitar duplicar y desactualizar una lista extensa de campos.

## 6. Explicación conceptual detallada por archivo

Las siguientes secciones usan un patrón completo para cada dato: **qué representa**, **interpretación**, **fórmula conceptual** y, cuando aplica, **fórmula implementada**. La fuente técnica exacta se conserva en las tablas anteriores.

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

La hoja `afiliados` del archivo trimestral corresponde al número de afiliados distribuido por administradora y tipo de fondo. Desde la migración a Teradata, estos valores no dependen del Excel local del Formato 491; se consultan desde `PROD_DWH_CONSULTA.FORMATO491` con `RENGLON = 999`, códigos de entidad de AFP y reglas de unidad de captura/tipo de fondo equivalentes a la hoja `multifondos`.

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

Los afiliados menores de 30 provienen de query Teradata sobre `PROD_DWH_CONSULTA.FORMATO491`, con las reglas de subcuenta y unidad de captura documentadas para el rango de edad.

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

El numerador proviene de query Teradata sobre `PROD_DWH_CONSULTA.FORMATO491`, usando las subcuentas/unidades de captura parametrizadas para el rango de 30 a 44 años.

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

El numerador proviene de query Teradata sobre `PROD_DWH_CONSULTA.FORMATO491`, usando las subcuentas/unidades de captura parametrizadas para el rango de 45 a 59 años.

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

El numerador proviene de query Teradata sobre `PROD_DWH_CONSULTA.FORMATO491`, usando las subcuentas/unidades de captura parametrizadas para mayores de 60 años.

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

#### Fila 15: Salario promedio en USD

**¿Qué representa la fila 15?**

La fila 15 corresponde al salario promedio ponderado de los afiliados cotizantes, convertido a dólares.

**Interpretación económica u operativa**

Permite aproximar y comparar internacionalmente el nivel salarial promedio sobre el que se realizan aportes.

**Fórmula conceptual**

$$
\text{Salario promedio en USD} = \frac{\text{Salario promedio ponderado en COP}}{\text{TRM}}
$$

**Fórmula implementada**

$$
\text{Fila 15} = \text{mensual.smColombiaUsd()}
$$

La aplicación calcula el salario ponderado en COP mediante la query del Formato 491 y los salarios mínimos oficiales de `SalarioMinimo.csv`; después lo convierte con la TRM del corte.

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
\text{Fila 16} = \text{Q495.total()}
$$

La aplicación consulta `PROD_DWH_CONSULTA.S9_FORMATO_495` para la fecha de corte, `UNIDAD_CAPTURA=1` y `RENGLON=200`. La fórmula de la macro Excel y la query vigente se detallan en la matriz 4.2.

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
\text{Fila 17} = \frac{\text{Q495.invalidez()}}{\text{Fila 16}}
$$

El numerador y el total se obtienen una sola vez de la query Q495. La celda se presenta con formato porcentual.

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
\text{Fila 18} = \frac{\text{Q495.vejez()}}{\text{Fila 16}}
$$

El numerador y el total se obtienen una sola vez de la query Q495. La celda se presenta con formato porcentual.

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
\text{Fila 19} = \frac{\text{Q495.sobrevivencia()}}{\text{Fila 16}}
$$

El numerador y el total se obtienen una sola vez de la query Q495. La celda se presenta con formato porcentual.

#### Fila 20: Altas de beneficiarios por tipo de prestación

**¿Qué representa la fila 20?**

La fila 20 identifica el bloque destinado a informar las nuevas altas de beneficiarios, clasificadas por tipo de prestación.

**Interpretación económica u operativa**

Permitiría medir el ingreso de nuevos beneficiarios al sistema durante el período. En la versión vigente no existe una fuente implementada para este bloque.

**Fórmula conceptual**

$
\text{Altas de beneficiarios} = \text{Nuevos beneficiarios reconocidos en el período}
$

**Fórmula implementada**

$
\text{Fila 20} = \text{“no disponible”}
$

La macro histórica escribe \`no disponible\` en esta fila. La aplicación vigente conserva el bloque sin calcular las filas 21–24.

#### Fila 21: Total de altas de beneficiarios

**¿Qué representa la fila 21?**

La fila 21 está destinada al total de nuevas altas de beneficiarios, sin distinguir el tipo de prestación.

**Interpretación económica u operativa**

Permitiría dimensionar el flujo total de nuevos beneficiarios del período.

**Fórmula conceptual**

$
\text{Total de altas} = \text{Altas por vejez} + \text{Altas por invalidez} + \text{Altas por sobrevivencia}
$

**Fórmula implementada**

No existe cálculo implementado para la fila 21; la salida conserva la celda de la plantilla sin poblarla.

#### Fila 22: Altas por vejez (%)

**¿Qué representa la fila 22?**

La fila 22 está destinada a la participación de las altas originadas por prestaciones de vejez.

**Interpretación económica u operativa**

Permitiría conocer qué proporción de los nuevos beneficiarios corresponde a pensiones de vejez.

**Fórmula conceptual**

$
\text{Altas por vejez (\%)} = \frac{\text{Altas por vejez}}{\text{Total de altas}} \times 100
$

**Fórmula implementada**

No existe cálculo implementado para la fila 22; la salida conserva la celda de la plantilla sin poblarla.

#### Fila 23: Altas por invalidez (%)

**¿Qué representa la fila 23?**

La fila 23 está destinada a la participación de las altas originadas por prestaciones de invalidez.

**Interpretación económica u operativa**

Permitiría conocer qué proporción de los nuevos beneficiarios corresponde a pensiones de invalidez.

**Fórmula conceptual**

$
\text{Altas por invalidez (\%)} = \frac{\text{Altas por invalidez}}{\text{Total de altas}} \times 100
$

**Fórmula implementada**

No existe cálculo implementado para la fila 23; la salida conserva la celda de la plantilla sin poblarla.

#### Fila 24: Altas por sobrevivencia (%)

**¿Qué representa la fila 24?**

La fila 24 está destinada a la participación de las altas originadas por prestaciones de sobrevivencia.

**Interpretación económica u operativa**

Permitiría conocer qué proporción de los nuevos beneficiarios corresponde a prestaciones de sobrevivencia.

**Fórmula conceptual**

$
\text{Altas por sobrevivencia (\%)} = \frac{\text{Altas por sobrevivencia}}{\text{Total de altas}} \times 100
$

**Fórmula implementada**

No existe cálculo implementado para la fila 24; la salida conserva la celda de la plantilla sin poblarla.

#### Fila 25: Movimiento de afiliados desde Formato 493 (miles)

**¿Qué representa la fila 25?**

La fila 25 corresponde al total de afiliados fallecidos del sistema durante la ventana anual terminada en la fecha de corte, expresado en miles.

**Interpretación económica u operativa**

Mide las salidas del sistema por fallecimiento y permite comparar su magnitud con otros flujos de afiliados.

**Fórmula conceptual**

$$
\text{Fallecimientos en miles} = \frac{\text{Afiliados fallecidos en la ventana anual}}{1000}
$$

**Fórmula implementada**

$$
\text{Fila 25} = \frac{\text{Q493.fallecidosSistema()}}{1000}
$$

La aplicación consulta `PROD_DWH_CONSULTA.S9_FORMATO_493` con `UNIDAD_CAPTURA=1`, renglones 165, 170 y 175, y la ventana anual aplicable. La fórmula histórica de Excel y la query vigente se detallan en la matriz 4.2.

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

Proviene de la query Q493 de traspasos para la ventana anual; el mismo agregado se reutiliza en la fila 27.

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
\text{Fila 28} = \frac{\text{mensual.vrFondo()}}{\text{TRM}}
$$

`mensual.vrFondo()` proviene de la query `ESTFIN_INDIV_PA`, PUC `100000`, ya expresada en millones de COP. La TRM se consulta una vez para el corte y se reutiliza.

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

#### Fila 32: Deuda de instituciones financieras locales (%)

**¿Qué representa la fila 32?**

La fila 32 corresponde a la proporción del portafolio total invertida en deuda emitida por instituciones financieras locales.

**Interpretación económica u operativa**

Permite identificar la exposición del portafolio al sector financiero local y evaluar su concentración por tipo de emisor.

**Fórmula conceptual**

$$
\text{Deuda de instituciones financieras locales} = \frac{\text{Valor invertido en deuda de instituciones financieras locales}}{\text{Valor total del portafolio}} \times 100
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

#### Fila 35: Administradores de fondos locales (%)

**¿Qué representa la fila 35?**

La fila 35 corresponde a la proporción del portafolio total invertida mediante administradores de fondos locales.

**Interpretación económica u operativa**

Permite identificar la exposición a vehículos gestionados por administradores de fondos locales.

**Fórmula conceptual**

$$
\text{Administradores de fondos locales} = \frac{\text{Valor invertido mediante administradores de fondos locales}}{\text{Valor total del portafolio}} \times 100
$$

**Fórmula implementada**

$$
\text{Fila 35} = \text{LIMITES!AIOS!K4}
$$

El valor proviene del archivo `LIMITES`, hoja `AIOS`, o del campo mensual equivalente para esta categoría.

#### Fila 36: Sociedades titulizadoras locales (%)

**¿Qué representa la fila 36?**

La fila 36 corresponde a la participación de inversiones locales en sociedades titulizadoras.

**Interpretación económica u operativa**

Permite identificar esta categoría de inversión; en la implementación vigente no se reporta exposición y se escribe cero.

**Fórmula conceptual**

$$
\text{Sociedades titulizadoras locales} = \frac{\text{Valor invertido en sociedades titulizadoras locales}}{\text{Valor total del portafolio}} \times 100
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
\text{Fila 37} = \text{LIMITES!AIOS!O4}
$$

El valor proviene del archivo `LIMITES`, hoja `AIOS`, o del campo mensual equivalente para esta categoría.

#### Fila 38: Deuda de instituciones financieras del exterior (%)

**¿Qué representa la fila 38?**

La fila 38 corresponde a la proporción del portafolio total invertida en deuda emitida por instituciones financieras del exterior.

**Interpretación económica u operativa**

Permite identificar la exposición del portafolio al sector financiero internacional.

**Fórmula conceptual**

$$
\text{Deuda de instituciones financieras del exterior} = \frac{\text{Valor invertido en deuda de instituciones financieras del exterior}}{\text{Valor total del portafolio}} \times 100
$$

**Fórmula implementada**

$$
\text{Fila 38} = \text{LIMITES!AIOS!Q4}
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
\text{Fila 39} = \text{LIMITES!AIOS!S4}
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
\text{Fila 40} = \text{LIMITES!AIOS!U4}
$$

El valor proviene del archivo `LIMITES`, hoja `AIOS`, o del campo mensual equivalente para esta categoría.

#### Fila 41: Administradores de fondos del exterior (%)

**¿Qué representa la fila 41?**

La fila 41 corresponde a la proporción del portafolio total invertida mediante administradores de fondos del exterior.

**Interpretación económica u operativa**

Permite identificar la exposición a vehículos gestionados por administradores internacionales.

**Fórmula conceptual**

$$
\text{Administradores de fondos del exterior} = \frac{\text{Valor invertido mediante administradores de fondos del exterior}}{\text{Valor total del portafolio}} \times 100
$$

**Fórmula implementada**

$$
\text{Fila 41} = \text{LIMITES!AIOS!W4}
$$

El valor proviene del archivo `LIMITES`, hoja `AIOS`, o del campo mensual equivalente para esta categoría.

#### Fila 42: Sociedades titulizadoras del exterior (%)

**¿Qué representa la fila 42?**

La fila 42 corresponde a la participación de inversiones del exterior en sociedades titulizadoras.

**Interpretación económica u operativa**

Permite identificar esta categoría de inversión exterior. La macro y la aplicación vigente usan fuentes distintas para poblarla.

**Fórmula conceptual**

$$
\text{Sociedades titulizadoras del exterior} = \frac{\text{Valor invertido en sociedades titulizadoras del exterior}}{\text{Valor total del portafolio}} \times 100
$$

**Fórmula implementada**

$$
\text{Fila 42} = 2
$$

La macro histórica toma `LIMITES!AIOS!Y4`; la aplicación vigente escribe la constante `2`. Esta diferencia se conserva explícita en la matriz 4.3.

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

#### Fila 46: Número de administradoras

**¿Qué representa la fila 46?**

La fila 46 corresponde al número de administradoras de fondos de pensiones incluidas en el informe.

**Interpretación económica u operativa**

Dimensiona la cantidad de entidades que participan en el sistema reportado.

**Fórmula conceptual**

$$
\text{Número de administradoras} = \text{Entidades incluidas en el informe}
$$

**Fórmula implementada**

$$
\text{Fila 46} = 4
$$

La aplicación escribe `4`: Colfondos, Porvenir, Protección y Skandia.

#### Fila 47: Participación Protección + Porvenir (%)

**¿Qué representa la fila 47?**

La fila 47 corresponde a la participación conjunta de Protección y Porvenir sobre el total del sistema.

**Interpretación económica u operativa**

Mide concentración de mercado de las administradoras seleccionadas dentro del total de fondos.

**Fórmula conceptual**

$$
\text{Participación Protección y Porvenir} = \frac{\text{Fondos Protección} + \text{Fondos Porvenir}}{\text{Fondos totales del sistema}} \times 100
$$

**Fórmula implementada**

$$
\text{Fila 47} = \frac{\text{Fondos Protección} + \text{Fondos Porvenir}}{\text{Fondos del sistema}}
$$

Los tres valores se consultan en `ESTFIN_INDIV_PA` para el PUC `100000`. La aplicación divide el porcentaje calculado por 100 antes de escribirlo porque la celda usa formato porcentual.

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
\text{Fila 61} = \frac{\text{Query Teradata Formato 136}/\text{TRM}}{\text{Aportantes}/1000} \times 1000
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

<!-- Fin del documento -->
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

## 7. Fuentes que continúan en archivos

| Fuente | Uso |
|---|---|
| `LIMITES del nuevo.xlsm` | Composición local/exterior y total de inversiones. |
| `PIB_PEA_TRM_DG` | PEA, PIB, deuda gubernamental y TRM de contingencia. |
| `Plantilla AIOS-probable.xlsm` | Cuentas, balances, resultados y gastos. Sus hojas `base anual` y `base mes` deben estar actualizadas antes de generar. |
| `Rent_Vr_Uni_Moderado.xlsm` | IPC y referencias para rentabilidad real. |
| `Valores_Fondo_Moder` / `MODERADO` | NAV histórico usado para rentabilidades nominales. |
| `Comisión FPO desde 2003.xlsx` | Comisiones obligatorias y de seguro por AFP. |

## 8. Controles de generación

- Los montos en COP se convierten a USD con la única TRM consultada para el período.
- Las divisiones usan `safeDivide`: si el denominador es cero, se escribe cero.
- El mensual conserva una fila por mes; el trimestral ordena los cierres cronológicamente y usa `sep`; el semestral usa una columna por junio o diciembre.
- Las constantes y los datos no disponibles se identifican expresamente; no se presentan como valores obtenidos de una fuente externa.
- La aplicación registra en el log la fuente y el cálculo de cada fila semestral para facilitar la auditoría.

<!-- Fin del documento -->
