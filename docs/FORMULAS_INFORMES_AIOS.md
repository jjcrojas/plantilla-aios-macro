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

## 6. Fuentes que continúan en archivos

| Fuente | Uso |
|---|---|
| `LIMITES del nuevo.xlsm` | Composición local/exterior y total de inversiones. |
| `PIB_PEA_TRM_DG` | PEA, PIB, deuda gubernamental y TRM de contingencia. |
| `Plantilla AIOS-probable.xlsm` | Cuentas, balances, resultados y gastos. Sus hojas `base anual` y `base mes` deben estar actualizadas antes de generar. |
| `Rent_Vr_Uni_Moderado.xlsm` | IPC y referencias para rentabilidad real. |
| `Valores_Fondo_Moder` / `MODERADO` | NAV histórico usado para rentabilidades nominales. |
| `Comisión FPO desde 2003.xlsx` | Comisiones obligatorias y de seguro por AFP. |

## 7. Controles de generación

- Los montos en COP se convierten a USD con la única TRM consultada para el período.
- Las divisiones usan `safeDivide`: si el denominador es cero, se escribe cero.
- El mensual conserva una fila por mes; el trimestral ordena los cierres cronológicamente y usa `sep`; el semestral usa una columna por junio o diciembre.
- Las constantes y los datos no disponibles se identifican expresamente; no se presentan como valores obtenidos de una fuente externa.
- La aplicación registra en el log la fuente y el cálculo de cada fila semestral para facilitar la auditoría.

<!-- Fin del documento -->
