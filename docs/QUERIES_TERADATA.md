--query 136_1
SELECT
fecha,
nombre_entidad,
Nombre_Fondo,
tipo_patrimonio,
codigo_patrimonio,
U_Captura,
"cod. renglón",
"nombre renglón",
SUM(Valor$) Valor$
FROM
(
SELECT   DISTINCT TRIM(YEAR(b.fecha))||'-'||TRIM(MONTH(b.fecha)) fecha,
         CASE 	
         		WHEN a.codigo_entidad = 2 THEN 'PROTECCIÓN'
         		WHEN a.codigo_entidad = 3 THEN 'PORVENIR'
                           WHEN a.codigo_entidad = 5 THEN 'HORIZONTE'
                           WHEN a.codigo_entidad = 8 THEN 'ING'
                           WHEN a.codigo_entidad = 9 THEN 'SKANDIA'
                            WHEN a.codigo_entidad = 10 THEN 'COLFONDOS'
                           
	END Nombre_Entidad,
         c.tipo_patrimonio,
         c.codigo_patrimonio,
         CASE 	
         		   WHEN c.tipo_patrimonio = 6 AND  c.codigo_patrimonio=1000 THEN 'PO_Moderado'
                           WHEN c.tipo_patrimonio = 6 AND  c.codigo_patrimonio=5000 THEN 'PO_Conservador'
                           WHEN c.tipo_patrimonio = 6 AND  c.codigo_patrimonio=6000 THEN 'PO_Mayor_Riesgo'
                           WHEN c.tipo_patrimonio = 6 AND  c.codigo_patrimonio=7000 THEN 'PO_Retiro_Programado'
                           WHEN c.tipo_patrimonio = 6 AND  c.codigo_patrimonio=8000 THEN 'PO_Skandia_Alternativo'
                           WHEN c.tipo_patrimonio = 5 AND  c.codigo_patrimonio=1 THEN 'CES_Largo_Plazo'
                           WHEN c.tipo_patrimonio = 5 AND  c.codigo_patrimonio=2 THEN 'CES_Corto_Plazo'
	END Nombre_Fondo,
         d.nivel3 "U_Captura",
         d.nivel4 "cod. renglón",
         d.desc_nivel4 "nombre renglón",
         SUM(e.valor) OVER (PARTITION BY fecha,a.codigo_entidad,c.tipo_patrimonio,c.codigo_patrimonio,d.nivel3,d.nivel4 ORDER BY Mes,c.tipo_patrimonio,c.codigo_patrimonio,d.nivel3,d.nivel4) Valor$
FROM     prod_dwh_consulta.entidades a,
         prod_dwh_consulta.tiempo b,
         prod_dwh_consulta.patrimonios_autonomos c,
         prod_dwh_consulta.negfid_insumos d,
         prod_dwh_consulta.negfid_insumo_entidad e
WHERE    d.inf_id=e.inf_id AND
         e.ent_id=a.ent_id AND
         e.tie_id=b.tie_id AND
         e.paau_id=c.paau_id AND
         c.tipo_patrimonio IN (5,6) AND
         c.codigo_patrimonio IN (1,2,1000,5000,6000,7000,8000) AND
         d.nivel1 IN (136) AND
         d.nivel2 = 2 AND 
         a.tipo_entidad = 23 AND
         e.valor<> 0 AND
         b.fecha > '2011-01-01' 
         ) mysub
         GROUP BY 1,2,3,4,5,6,7,8
         ORDER BY 1,2;
--query136_2
SELECT MAX(b.Fecha) max_fecha
FROM     prod_dwh_consulta.entidades a,
         prod_dwh_consulta.tiempo b,
         prod_dwh_consulta.patrimonios_autonomos c,
         prod_dwh_consulta.negfid_insumos d,
         prod_dwh_consulta.negfid_insumo_entidad e
WHERE    d.inf_id=e.inf_id AND
         e.ent_id=a.ent_id AND
         e.tie_id=b.tie_id AND
         e.paau_id=c.paau_id AND
         c.tipo_patrimonio IN (5,6) AND
         c.codigo_patrimonio IN (1,2,1000,5000,6000,7000,8000) AND
         d.nivel1 IN (136) AND
         d.nivel2 = 2 AND 
         a.tipo_entidad = 23 AND
         b.fecha >= '2011-01-01' 
--query136_3
SELECT 	b.fecha,
         		a.tipo_entidad,
         		a.nombre_tipo_entidad, 
         		a.codigo_entidad, 
         		 CASE 	
         		WHEN a.codigo_entidad = 2 THEN 'PROTECCIÓN'
         		WHEN a.codigo_entidad = 3 THEN 'PORVENIR'
                           WHEN a.codigo_entidad = 5 THEN 'HORIZONTE'
                           WHEN a.codigo_entidad = 8 THEN 'ING'
                           WHEN a.codigo_entidad = 9 THEN 'SKANDIA'
                            WHEN a.codigo_entidad = 10 THEN 'COLFONDOS'
                           
	END Nombre_Entidad,
         		c.tipo_patrimonio,
         		c.nombre_tipo_patrimonio,
         		c.subtipo_patrimonio,
         		c.nombre_subtipo_patrimonio,
         		c.codigo_patrimonio,
         		c.nombre_patrimonio,
         		d.Nivel1 "Formato",
                           d.desc_nivel1 "nombre formato",
         		d.nivel2 "codigo columna",
         		d.desc_nivel2 "nombre columna",
         		d.nivel3 "cod. unid. capt.",
         		d.desc_nivel3 "nombre unid.capt.",
         		d.nivel4 "cod. renglón",
         		d.desc_nivel4 "nombre renglón",
         		e.valor
FROM		prod_dwh_consulta.entidades a,
         		prod_dwh_consulta.tiempo b,
         		prod_dwh_consulta.patrimonios_autonomos c,
         		prod_dwh_consulta.negfid_insumos d,
         		prod_dwh_consulta.negfid_insumo_entidad e
WHERE		d.inf_id=e.inf_id AND
         		e.ent_id=a.ent_id AND
         		e.tie_id=b.tie_id AND
         		e.paau_id=c.paau_id AND
                           d.Tipo_Informe= 17 AND
         		d.nivel1 IN (136) AND
                           d.nivel3 IN (3,4) AND
                           d.nivel4 IN (5,105,110,300,305) AND
         		b.fecha <= Current_Date AND
                           b.fecha = Last_Day(b.fecha)
--query491
-- TIPO DE FONDO: MODERADO, CONSERVADOR Y DE MAYOR RIESGO 

SELECT            TIPO_ENTIDAD, 
                            CODIGO_ENTIDAD, 
                            NOMBRE_ENTIDAD, 
                            CIDT, 
                            FECBAL,
                            TIPO_IDENTIFICACION, 
                            NUMERO_IDENTIFICACION,
                            CASE WHEN SUBSTR ( NUMERO_IDENTIFICACION, 9 , 4 ) IN ('1000') THEN 'MODERADO'
                                        WHEN SUBSTR ( NUMERO_IDENTIFICACION, 9 , 4 ) IN ('5000') THEN 'CONSERVADOR'
                                        WHEN SUBSTR ( NUMERO_IDENTIFICACION, 9 , 4 ) IN ('6000') THEN 'MAYOR RIESGO'
                                        WHEN SUBSTR ( NUMERO_IDENTIFICACION, 9 , 4 ) IN ('7000') THEN 'RETIRO PROGRAMADO'
                                        WHEN SUBSTR ( NUMERO_IDENTIFICACION, 9 , 4 ) IN ('8000') THEN 'OM_ALTERNATIVO'
                            END Tipo_de_Fondo,            
                            UNIDAD_CAPTURA, 
                            NOMBRE_UNIDAD_CAPTURA,
                            RENGLON, 
                            CASE WHEN RENGLON = 5 THEN '15 años'
                                        WHEN RENGLON = 10 THEN '16 años'
                                        WHEN RENGLON = 15 THEN '17 años'
                                        WHEN RENGLON = 20 THEN '18 años'
                                        WHEN RENGLON = 25 THEN '19 años'
                                        WHEN RENGLON = 30 THEN '20 años'
                                        WHEN RENGLON = 35 THEN '21 años'
                                        WHEN RENGLON = 40 THEN '22 años'
                                        WHEN RENGLON = 45 THEN '23 años'
                                        WHEN RENGLON = 50 THEN '24 años'
                                        WHEN RENGLON = 55 THEN '25 años'
                                        WHEN RENGLON = 60 THEN '26 años'
                                        WHEN RENGLON = 65 THEN '27 años'
                                        WHEN RENGLON = 70 THEN '28 años'
                                        WHEN RENGLON = 75 THEN '29 años'
                                        WHEN RENGLON = 80 THEN '30 años'
                                        WHEN RENGLON = 85 THEN '31 años'
                                        WHEN RENGLON = 90 THEN '32 años'
                                        WHEN RENGLON = 95 THEN '33 años'
                                        WHEN RENGLON = 100 THEN '34 años'
                                        WHEN RENGLON = 105 THEN '35 años'
                                        WHEN RENGLON = 110 THEN '36 años'
                                        WHEN RENGLON = 115 THEN '37 años'
                                        WHEN RENGLON = 120 THEN '38 años'
                                        WHEN RENGLON = 125 THEN '39 años'
                                        WHEN RENGLON = 130 THEN '40 años'
                                        WHEN RENGLON = 135 THEN '41 años'
                                        WHEN RENGLON = 140 THEN '42 años'
                                        WHEN RENGLON = 145 THEN '43 años'
                                        WHEN RENGLON = 150 THEN '44 años'
                                        WHEN RENGLON = 155 THEN '45 años'
                                        WHEN RENGLON = 160 THEN '46 años'
                                        WHEN RENGLON = 165 THEN '47 años'
                                        WHEN RENGLON = 170 THEN '48 años'
                                        WHEN RENGLON = 175 THEN '49 años'
                                        WHEN RENGLON = 180 THEN '50 años'
                                        WHEN RENGLON = 185 THEN '51 años'
                                        WHEN RENGLON = 190 THEN '52 años'
                                        WHEN RENGLON = 195 THEN '53 años'
                                        WHEN RENGLON = 200 THEN '54 años'
                                        WHEN RENGLON = 205 THEN '55 años'
                                        WHEN RENGLON = 210 THEN '56 años'
                                        WHEN RENGLON = 215 THEN '57 años'
                                        WHEN RENGLON = 220 THEN '58 años'
                                        WHEN RENGLON = 225 THEN '59 años'
                                        WHEN RENGLON = 230 THEN '60 años'
                                        WHEN RENGLON = 235 THEN '61 años'
                                        WHEN RENGLON = 240 THEN '62 años'
                                        WHEN RENGLON = 245 THEN '63 años'
                                        WHEN RENGLON = 250 THEN '64 años'
                                        WHEN RENGLON = 255 THEN '65 o mas años'
                                        WHEN RENGLON = 999 THEN 'Total'
                            END Edades,            
                            SUM(TOTAL_AFILIADOS_H_1), 
                            SUM(PROMEDIO_SEMANAS_COTIZADAS_H_1),
                            SUM(PROM_SLD_CTA_INDIV_AHO_PEN_H_1), 
                            SUM(NO_AFILIADOS_ACTIVOS_H_1), 
                            SUM(NO_AFILIADOS_COTIZANTES_H_1),
                            SUM(NO_AFLDS_COBERT_SEG_PREV_H_1), 
                            SUM(TOTAL_AFILIADOS_M_1), 
                            SUM(PROMEDIO_SEMANAS_COTIZADAS_M_1),
                            SUM(PROM_SLD_CTA_INDIV_AHO_PEN_M_1), 
                            SUM(NO_AFILIADOS_ACTIVOS_M_1), 
                            SUM(NO_AFILIADOS_COTIZANTES_M_1),
                            SUM(NO_AFLDS_COBERT_SEG_PREV_M_1), 
                            SUM(TOTAL_AFILIADOS_H_1_2), 
                            SUM( PROM_SEMANAS_COTIZADAS_H_1_2),
                            SUM(PROM_SLD_CTA_IND_AHO_PEN_H_1_2), 
                            SUM(NO_AFILIADOS_ACTIVOS_H_1_2), 
                            SUM(NO_AFILIADOS_COTIZANTES_H_1_2),
                            SUM(NO_AFLDS_COBERT_SEG_PREV_H_1_2), 
                            SUM(TOTAL_AFILIADOS_M_1_2), 
                            SUM(PROM_SEMANAS_COTIZADAS_M_1_2),
                            SUM(PROM_SLD_CTA_IND_AHO_PEN_M_1_2), 
                            SUM(NO_AFILIADOS_ACTIVOS_M_1_2), 
                            SUM(NO_AFILIADOS_COTIZANTES_M_1_2),
                            SUM(NO_AFLDS_COBERT_SEG_PREV_M_1_2), 
                            SUM(TOTAL_AFILIADOS_H_2_3), 
                            SUM(PROM_SEMANAS_COTIZADAS_H_2_3),
                            SUM(PROM_SLD_CTA_IND_AHO_PEN_H_2_3), 
                            SUM(NO_AFILIADOS_ACTIVOS_H_2_3), 
                            SUM(NO_AFILIADOS_COTIZANTES_H_2_3),
                            SUM(NO_AFLDS_COBERT_SEG_PREV_H_2_3), 
                            SUM(TOTAL_AFILIADOS_M_2_3), 
                            SUM(PROM_SEMANAS_COTIZADAS_M_2_3),
                            SUM(PROM_SLD_CTA_IND_AHO_PEN_M_2_3), 
                            SUM(NO_AFILIADOS_ACTIVOS_M_2_3), 
                            SUM(NO_AFILIADOS_COTIZANTES_M_2_3),
                            SUM(NO_AFLDS_COBERT_SEG_PREV_M_2_3), 
                            SUM(TOTAL_AFILIADOS_H_3_4), 
                            SUM(PROM_SEMANAS_COTIZADAS_H_3_4),
                            SUM(PROM_SLD_CTA_IND_AHO_PEN_H_3_4), 
                            SUM(NO_AFILIADOS_ACTIVOS_H_3_4), 
                            SUM(NO_AFILIADOS_COTIZANTES_H_3_4),
                            SUM(NO_AFLDS_COBERT_SEG_PREV_H_3_4), 
                            SUM(TOTAL_AFILIADOS_M_3_4), 
                            SUM(PROM_SEMANAS_COTIZADAS_M_3_4),
                            SUM(PROM_SLD_CTA_IND_AHO_PEN_M_3_4), 
                            SUM(NO_AFILIADOS_ACTIVOS_M_3_4), 
                            SUM(NO_AFILIADOS_COTIZANTES_M_3_4),
                            SUM(NO_AFLDS_COBERT_SEG_PREV_M_3_4), 
                            SUM(TOTAL_AFILIADOS_H_4_8), 
                            SUM(PROM_SEMANAS_COTIZADAS_H_4_8),
                            SUM(PROM_SLD_CTA_IND_AHO_PEN_H_4_8), 
                            SUM(NO_AFILIADOS_ACTIVOS_H_4_8), 
                            SUM(NO_AFILIADOS_COTIZANTES_H_4_8),
                            SUM(NO_AFLDS_COBERT_SEG_PREV_H_4_8), 
                            SUM(TOTAL_AFILIADOS_M_4_8), 
                            SUM(PROM_SEMANAS_COTIZADAS_M_4_8),
                            SUM(PROM_SLD_CTA_IND_AHO_PEN_M_4_8), 
                            SUM(NO_AFILIADOS_ACTIVOS_M_4_8), 
                            SUM(NO_AFILIADOS_COTIZANTES_M_4_8),
                            SUM( NO_AFLDS_COBERT_SEG_PREV_M_4_8), 
                            SUM(TOTAL_AFILIADOS_H_8_12), 
                            SUM(PROM_SEMANAS_COTIZADAS_H_8_12),
                            SUM(PRM_SLD_CTA_IND_AHO_PEN_H_8_12), 
                            SUM(NO_AFILIADOS_ACTIVOS_H_8_12),
                            SUM(NO_AFILIADOS_COTIZANTES_H_8_12), 
                            SUM(NO_AFLDS_COBER_SEG_PREV_H_8_12),
                            SUM(TOTAL_AFILIADOS_M_8_12), 
                            SUM(PROM_SEMANAS_COTIZADAS_M_8_12), 
                            SUM(PRM_SLD_CTA_IND_AHO_PEN_M_8_12),
                            SUM(NO_AFILIADOS_ACTIVOS_M_8_12), 
                            SUM(NO_AFILIADOS_COTIZANTES_M_8_12),
                            SUM(NO_AFLDS_COBER_SEG_PREV_M_8_12), 
                            SUM(TOTAL_AFILIADOS_H_12_16), 
                            SUM(PROM_SEMANAS_COTIZADAS_H_12_16),
                            SUM(PRM_SLD_CTA_IND_AHOPEN_H_12_16), 
                            SUM(NO_AFILIADOS_ACTIVOS_H_12_16),
                            SUM(NO_AFLDS_COTIZANTES_H_12_16), 
                            SUM(NO_AFLDS_COBER_SEG_PRE_H_12_16),
                            SUM(TOTAL_AFILIADOS_M_12_16), 
                            SUM(PROM_SEMANAS_COTIZADAS_M_12_16), 
                            SUM(PRM_SLD_CTA_IND_AHOPEN_M_12_16),
                            SUM(NO_AFILIADOS_ACTIVOS_M_12_16), 
                            SUM(NO_AFLDS_COTIZANTES_M_12_16), 
                            SUM(NO_AFLDS_COBER_SEG_PRE_M_12_16),
                            SUM(TOTAL_AFILIADOS_H_16_20), 
                            SUM(PROM_SEMANAS_COTIZADAS_H_16_20), 
                            SUM(PRM_SLD_CTA_IND_AHOPEN_H_16_20),
                            SUM(NO_AFILIADOS_ACTIVOS_H_16_20), 
                            SUM(NO_AFLDS_COTIZANTES_H_16_20), 
                            SUM(NO_AFLDS_COBER_SEG_PRE_H_16_20),
                            SUM(TOTAL_AFILIADOS_M_16_20), 
                            SUM(PROM_SEMANAS_COTIZADAS_M_16_20), 
                            SUM(PRM_SLD_CTA_IND_AHOPEN_M_16_20),
                            SUM(NO_AFILIADOS_ACTIVOS_M_16_20),
                            SUM(NO_AFLDS_COTIZANTES_M_16_20), 
                            SUM(NO_AFLDS_COBER_SEG_PRE_M_16_20),
                            SUM(TOTAL_AFILIADOS_H_20), 
                            SUM(PROM_SEMANAS_COTIZADAS_H_20), 
                            SUM(PROM_SLD_CTA_IND_AHO_PEN_H_20),
                            SUM(NO_AFILIADOS_ACTIVOS_H_20), 
                            SUM(NO_AFILIADOS_COTIZANTES_H_20), 
                            SUM(NO_AFLDS_COBERT_SEG_PREV_H_20),
                            SUM(TOTAL_AFILIADOS_M_20), 
                            SUM(PROM_SEMANAS_COTIZADAS_M_20),
                            SUM(PROM_SLD_CTA_IND_AHO_PEN_M_20),
                            SUM(NO_AFILIADOS_ACTIVOS_M_20), 
                            SUM(NO_AFILIADOS_COTIZANTES_M_20), 
                            SUM(NO_AFLDS_COBERT_SEG_PREV_M_20),
                            SUM(TOTAL_AFILIADOS_H), 
                            SUM(TOTAL_AFILIADOS_M), 
                            SUM(TOTAL_AFILIADOS_TOTAL),
                            SUM(TOTAL_AFILIADOS_ACTIVOS_H), 
                            SUM(TOTAL_AFILIADOS_ACTIVOS_M), 
                            SUM(TOTAL_AFILIADOS_ACTIVOS_TOTAL),
                            SUM(TOTAL_AFILIADOS_COTIZANTES_H), 
                            SUM(TOTAL_AFILIADOS_COTIZANTES_M), 
                            SUM(TOTAL_AFILIADOS_COTIZANTES),
                            SUM(TIPO_AFILIACION_DEPENDIENTE), 
                            SUM(TIPO_AFILIACION_INDEPENDIENTE), 
                            SUM(ORIGEN_DE_AFILIACION_ISS),
                            SUM(ORIGEN_DE_AFILIACION_CAJAS), 
                            SUM(ORIGEN_AFILIACION_INGRESO), 
                            SUM(ORIGEN_AFILIACION_TRASLADO)
FROM                PROD_DWH_CONSULTA.FORMATO491
WHERE             UNIDAD_CAPTURA =1 
							--FECBAL BETWEEN '2019-01-12' AND '2020-03-31'
							--FECBAL = '2020-01-31' 
GROUP BY      TIPO_ENTIDAD, 
                            CODIGO_ENTIDAD, 
                            NOMBRE_ENTIDAD, 
                            CIDT, 
                            FECBAL,
                            TIPO_IDENTIFICACION, 
                            NUMERO_IDENTIFICACION,
                             UNIDAD_CAPTURA, 
                            NOMBRE_UNIDAD_CAPTURA,
                            RENGLON
UNION
--  TIPO DE FONDO CONSERVADOR (CONVERGENCIA CONSERVADOR Y MODERADO) 
SELECT            TIPO_ENTIDAD, 
                            CODIGO_ENTIDAD, 
                            NOMBRE_ENTIDAD, 
                            CIDT, 
                            FECBAL,
                            TIPO_IDENTIFICACION, 
                            NUMERO_IDENTIFICACION,
                            'CONS Y MODER' Tipo_de_Fondo,            
                            UNIDAD_CAPTURA, 
                            NOMBRE_UNIDAD_CAPTURA,
                            RENGLON, 
                            CASE WHEN RENGLON = 5 THEN '50 años'
                                        WHEN RENGLON = 10 THEN '51 años'
                                        WHEN RENGLON = 15 THEN '52 años'
                                        WHEN RENGLON = 20 THEN '53 años'
                                        WHEN RENGLON = 25 THEN '54 años'
                                        WHEN RENGLON = 30 THEN '55 años'
                                        WHEN RENGLON = 35 THEN '56 años'
                                        WHEN RENGLON = 40 THEN '57 años'
                                        WHEN RENGLON = 45 THEN '58 años'
                                        WHEN RENGLON = 50 THEN '59 años'
                                        WHEN RENGLON = 55 THEN '60 años'
                                        WHEN RENGLON = 60 THEN '61 años'
                                        WHEN RENGLON = 65 THEN '62 años'
                                        WHEN RENGLON = 70 THEN '63 años'
                                        WHEN RENGLON = 75 THEN '64 años'
                                        WHEN RENGLON = 80 THEN '65 años'
                                        WHEN RENGLON = 999 THEN 'Total'
                            END Edades,
                            SUM(TOTAL_AFILIADOS_H_1), 
                            SUM(PROMEDIO_SEMANAS_COTIZADAS_H_1),
                            SUM(PROM_SLD_CTA_INDIV_AHO_PEN_H_1), 
                            SUM(NO_AFILIADOS_ACTIVOS_H_1), 
                            SUM(NO_AFILIADOS_COTIZANTES_H_1),
                            SUM(NO_AFLDS_COBERT_SEG_PREV_H_1), 
                            SUM(TOTAL_AFILIADOS_M_1), 
                            SUM(PROMEDIO_SEMANAS_COTIZADAS_M_1),
                            SUM(PROM_SLD_CTA_INDIV_AHO_PEN_M_1), 
                            SUM(NO_AFILIADOS_ACTIVOS_M_1), 
                            SUM(NO_AFILIADOS_COTIZANTES_M_1),
                            SUM(NO_AFLDS_COBERT_SEG_PREV_M_1), 
                            SUM(TOTAL_AFILIADOS_H_1_2), 
                            SUM( PROM_SEMANAS_COTIZADAS_H_1_2),
                            SUM(PROM_SLD_CTA_IND_AHO_PEN_H_1_2), 
                            SUM(NO_AFILIADOS_ACTIVOS_H_1_2), 
                            SUM(NO_AFILIADOS_COTIZANTES_H_1_2),
                            SUM(NO_AFLDS_COBERT_SEG_PREV_H_1_2), 
                            SUM(TOTAL_AFILIADOS_M_1_2), 
                            SUM(PROM_SEMANAS_COTIZADAS_M_1_2),
                            SUM(PROM_SLD_CTA_IND_AHO_PEN_M_1_2), 
                            SUM(NO_AFILIADOS_ACTIVOS_M_1_2), 
                            SUM(NO_AFILIADOS_COTIZANTES_M_1_2),
                            SUM(NO_AFLDS_COBERT_SEG_PREV_M_1_2), 
                            SUM(TOTAL_AFILIADOS_H_2_3), 
                            SUM(PROM_SEMANAS_COTIZADAS_H_2_3),
                            SUM(PROM_SLD_CTA_IND_AHO_PEN_H_2_3), 
                            SUM(NO_AFILIADOS_ACTIVOS_H_2_3), 
                            SUM(NO_AFILIADOS_COTIZANTES_H_2_3),
                            SUM(NO_AFLDS_COBERT_SEG_PREV_H_2_3), 
                            SUM(TOTAL_AFILIADOS_M_2_3), 
                            SUM(PROM_SEMANAS_COTIZADAS_M_2_3),
                            SUM(PROM_SLD_CTA_IND_AHO_PEN_M_2_3), 
                            SUM(NO_AFILIADOS_ACTIVOS_M_2_3), 
                            SUM(NO_AFILIADOS_COTIZANTES_M_2_3),
                            SUM(NO_AFLDS_COBERT_SEG_PREV_M_2_3), 
                            SUM(TOTAL_AFILIADOS_H_3_4), 
                            SUM(PROM_SEMANAS_COTIZADAS_H_3_4),
                            SUM(PROM_SLD_CTA_IND_AHO_PEN_H_3_4), 
                            SUM(NO_AFILIADOS_ACTIVOS_H_3_4), 
                            SUM(NO_AFILIADOS_COTIZANTES_H_3_4),
                            SUM(NO_AFLDS_COBERT_SEG_PREV_H_3_4), 
                            SUM(TOTAL_AFILIADOS_M_3_4), 
                            SUM(PROM_SEMANAS_COTIZADAS_M_3_4),
                            SUM(PROM_SLD_CTA_IND_AHO_PEN_M_3_4), 
                            SUM(NO_AFILIADOS_ACTIVOS_M_3_4), 
                            SUM(NO_AFILIADOS_COTIZANTES_M_3_4),
                            SUM(NO_AFLDS_COBERT_SEG_PREV_M_3_4), 
                            SUM(TOTAL_AFILIADOS_H_4_8), 
                            SUM(PROM_SEMANAS_COTIZADAS_H_4_8),
                            SUM(PROM_SLD_CTA_IND_AHO_PEN_H_4_8), 
                            SUM(NO_AFILIADOS_ACTIVOS_H_4_8), 
                            SUM(NO_AFILIADOS_COTIZANTES_H_4_8),
                            SUM(NO_AFLDS_COBERT_SEG_PREV_H_4_8), 
                            SUM(TOTAL_AFILIADOS_M_4_8), 
                            SUM(PROM_SEMANAS_COTIZADAS_M_4_8),
                            SUM(PROM_SLD_CTA_IND_AHO_PEN_M_4_8), 
                            SUM(NO_AFILIADOS_ACTIVOS_M_4_8), 
                            SUM(NO_AFILIADOS_COTIZANTES_M_4_8),
                            SUM( NO_AFLDS_COBERT_SEG_PREV_M_4_8), 
                            SUM(TOTAL_AFILIADOS_H_8_12), 
                            SUM(PROM_SEMANAS_COTIZADAS_H_8_12),
                            SUM(PRM_SLD_CTA_IND_AHO_PEN_H_8_12), 
                            SUM(NO_AFILIADOS_ACTIVOS_H_8_12),
                            SUM(NO_AFILIADOS_COTIZANTES_H_8_12), 
                            SUM(NO_AFLDS_COBER_SEG_PREV_H_8_12),
                            SUM(TOTAL_AFILIADOS_M_8_12), 
                            SUM(PROM_SEMANAS_COTIZADAS_M_8_12), 
                            SUM(PRM_SLD_CTA_IND_AHO_PEN_M_8_12),
                            SUM(NO_AFILIADOS_ACTIVOS_M_8_12), 
                            SUM(NO_AFILIADOS_COTIZANTES_M_8_12),
                            SUM(NO_AFLDS_COBER_SEG_PREV_M_8_12), 
                            SUM(TOTAL_AFILIADOS_H_12_16), 
                            SUM(PROM_SEMANAS_COTIZADAS_H_12_16),
                            SUM(PRM_SLD_CTA_IND_AHOPEN_H_12_16), 
                            SUM(NO_AFILIADOS_ACTIVOS_H_12_16),
                            SUM(NO_AFLDS_COTIZANTES_H_12_16), 
                            SUM(NO_AFLDS_COBER_SEG_PRE_H_12_16),
                            SUM(TOTAL_AFILIADOS_M_12_16), 
                            SUM(PROM_SEMANAS_COTIZADAS_M_12_16), 
                            SUM(PRM_SLD_CTA_IND_AHOPEN_M_12_16),
                            SUM(NO_AFILIADOS_ACTIVOS_M_12_16), 
                            SUM(NO_AFLDS_COTIZANTES_M_12_16), 
                            SUM(NO_AFLDS_COBER_SEG_PRE_M_12_16),
                            SUM(TOTAL_AFILIADOS_H_16_20), 
                            SUM(PROM_SEMANAS_COTIZADAS_H_16_20), 
                            SUM(PRM_SLD_CTA_IND_AHOPEN_H_16_20),
                            SUM(NO_AFILIADOS_ACTIVOS_H_16_20), 
                            SUM(NO_AFLDS_COTIZANTES_H_16_20), 
                            SUM(NO_AFLDS_COBER_SEG_PRE_H_16_20),
                            SUM(TOTAL_AFILIADOS_M_16_20), 
                            SUM(PROM_SEMANAS_COTIZADAS_M_16_20), 
                            SUM(PRM_SLD_CTA_IND_AHOPEN_M_16_20),
                            SUM(NO_AFILIADOS_ACTIVOS_M_16_20),
                            SUM(NO_AFLDS_COTIZANTES_M_16_20), 
                            SUM(NO_AFLDS_COBER_SEG_PRE_M_16_20),
                            SUM(TOTAL_AFILIADOS_H_20), 
                            SUM(PROM_SEMANAS_COTIZADAS_H_20), 
                            SUM(PROM_SLD_CTA_IND_AHO_PEN_H_20),
                            SUM(NO_AFILIADOS_ACTIVOS_H_20), 
                            SUM(NO_AFILIADOS_COTIZANTES_H_20), 
                            SUM(NO_AFLDS_COBERT_SEG_PREV_H_20),
                            SUM(TOTAL_AFILIADOS_M_20), 
                            SUM(PROM_SEMANAS_COTIZADAS_M_20),
                            SUM(PROM_SLD_CTA_IND_AHO_PEN_M_20),
                            SUM(NO_AFILIADOS_ACTIVOS_M_20), 
                            SUM(NO_AFILIADOS_COTIZANTES_M_20), 
                            SUM(NO_AFLDS_COBERT_SEG_PREV_M_20),
                            SUM(TOTAL_AFILIADOS_H), 
                            SUM(TOTAL_AFILIADOS_M), 
                            SUM(TOTAL_AFILIADOS_TOTAL),
                            SUM(TOTAL_AFILIADOS_ACTIVOS_H), 
                            SUM(TOTAL_AFILIADOS_ACTIVOS_M), 
                            SUM(TOTAL_AFILIADOS_ACTIVOS_TOTAL),
                            SUM(TOTAL_AFILIADOS_COTIZANTES_H), 
                            SUM(TOTAL_AFILIADOS_COTIZANTES_M), 
                            SUM(TOTAL_AFILIADOS_COTIZANTES),
                            SUM(TIPO_AFILIACION_DEPENDIENTE), 
                            SUM(TIPO_AFILIACION_INDEPENDIENTE), 
                            SUM(ORIGEN_DE_AFILIACION_ISS),
                            SUM(ORIGEN_DE_AFILIACION_CAJAS), 
                            SUM(ORIGEN_AFILIACION_INGRESO), 
                            SUM(ORIGEN_AFILIACION_TRASLADO)
FROM                PROD_DWH_CONSULTA.FORMATO491
WHERE            UNIDAD_CAPTURA =2 
                             ---FECBAL BETWEEN '2019-01-12' AND '2020-03-31'
                             ---FECBAL = '2020-01-31' 
GROUP BY      TIPO_ENTIDAD, 
                            CODIGO_ENTIDAD, 
                            NOMBRE_ENTIDAD, 
                            CIDT, 
                            FECBAL,
                            TIPO_IDENTIFICACION, 
                            NUMERO_IDENTIFICACION,
                             UNIDAD_CAPTURA, 
                            NOMBRE_UNIDAD_CAPTURA,
                            RENGLON
--ORDER BY   --    FECBAL,
                        --    CODIGO_ENTIDAD
UNION
--   TIPO DE FONDO CONSERVADOR (CONVERGENCIA CONSERVADOR Y DE MAYOR RIESGO) 
SELECT            TIPO_ENTIDAD, 
                            CODIGO_ENTIDAD, 
                            NOMBRE_ENTIDAD, 
                            CIDT, 
                            FECBAL,
                            TIPO_IDENTIFICACION, 
                            NUMERO_IDENTIFICACION,
                            'CONS Y MAY RIES' Tipo_de_Fondo,            
                            UNIDAD_CAPTURA, 
                            NOMBRE_UNIDAD_CAPTURA,
                            RENGLON, 
                            CASE WHEN RENGLON = 5 THEN '50 años'
                                        WHEN RENGLON = 10 THEN '51 años'
                                        WHEN RENGLON = 15 THEN '52 años'
                                        WHEN RENGLON = 20 THEN '53 años'
                                        WHEN RENGLON = 25 THEN '54 años'
                                        WHEN RENGLON = 30 THEN '55 años'
                                        WHEN RENGLON = 35 THEN '56 años'
                                        WHEN RENGLON = 40 THEN '57 años'
                                        WHEN RENGLON = 45 THEN '58 años'
                                        WHEN RENGLON = 50 THEN '59 años'
                                        WHEN RENGLON = 55 THEN '60 años'
                                        WHEN RENGLON = 60 THEN '61 años'
                                        WHEN RENGLON = 65 THEN '62 años'
                                        WHEN RENGLON = 70 THEN '63 años'
                                        WHEN RENGLON = 75 THEN '64 años'
                                        WHEN RENGLON = 80 THEN '65 años'
                                        WHEN RENGLON = 999 THEN 'Total'
                            END Edades,
                            SUM(TOTAL_AFILIADOS_H_1), 
                            SUM(PROMEDIO_SEMANAS_COTIZADAS_H_1),
                            SUM(PROM_SLD_CTA_INDIV_AHO_PEN_H_1), 
                            SUM(NO_AFILIADOS_ACTIVOS_H_1), 
                            SUM(NO_AFILIADOS_COTIZANTES_H_1),
                            SUM(NO_AFLDS_COBERT_SEG_PREV_H_1), 
                            SUM(TOTAL_AFILIADOS_M_1), 
                            SUM(PROMEDIO_SEMANAS_COTIZADAS_M_1),
                            SUM(PROM_SLD_CTA_INDIV_AHO_PEN_M_1), 
                            SUM(NO_AFILIADOS_ACTIVOS_M_1), 
                            SUM(NO_AFILIADOS_COTIZANTES_M_1),
                            SUM(NO_AFLDS_COBERT_SEG_PREV_M_1), 
                            SUM(TOTAL_AFILIADOS_H_1_2), 
                            SUM( PROM_SEMANAS_COTIZADAS_H_1_2),
                            SUM(PROM_SLD_CTA_IND_AHO_PEN_H_1_2), 
                            SUM(NO_AFILIADOS_ACTIVOS_H_1_2), 
                            SUM(NO_AFILIADOS_COTIZANTES_H_1_2),
                            SUM(NO_AFLDS_COBERT_SEG_PREV_H_1_2), 
                            SUM(TOTAL_AFILIADOS_M_1_2), 
                            SUM(PROM_SEMANAS_COTIZADAS_M_1_2),
                            SUM(PROM_SLD_CTA_IND_AHO_PEN_M_1_2), 
                            SUM(NO_AFILIADOS_ACTIVOS_M_1_2), 
                            SUM(NO_AFILIADOS_COTIZANTES_M_1_2),
                            SUM(NO_AFLDS_COBERT_SEG_PREV_M_1_2), 
                            SUM(TOTAL_AFILIADOS_H_2_3), 
                            SUM(PROM_SEMANAS_COTIZADAS_H_2_3),
                            SUM(PROM_SLD_CTA_IND_AHO_PEN_H_2_3), 
                            SUM(NO_AFILIADOS_ACTIVOS_H_2_3), 
                            SUM(NO_AFILIADOS_COTIZANTES_H_2_3),
                            SUM(NO_AFLDS_COBERT_SEG_PREV_H_2_3), 
                            SUM(TOTAL_AFILIADOS_M_2_3), 
                            SUM(PROM_SEMANAS_COTIZADAS_M_2_3),
                            SUM(PROM_SLD_CTA_IND_AHO_PEN_M_2_3), 
                            SUM(NO_AFILIADOS_ACTIVOS_M_2_3), 
                            SUM(NO_AFILIADOS_COTIZANTES_M_2_3),
                            SUM(NO_AFLDS_COBERT_SEG_PREV_M_2_3), 
                            SUM(TOTAL_AFILIADOS_H_3_4), 
                            SUM(PROM_SEMANAS_COTIZADAS_H_3_4),
                            SUM(PROM_SLD_CTA_IND_AHO_PEN_H_3_4), 
                            SUM(NO_AFILIADOS_ACTIVOS_H_3_4), 
                            SUM(NO_AFILIADOS_COTIZANTES_H_3_4),
                            SUM(NO_AFLDS_COBERT_SEG_PREV_H_3_4), 
                            SUM(TOTAL_AFILIADOS_M_3_4), 
                            SUM(PROM_SEMANAS_COTIZADAS_M_3_4),
                            SUM(PROM_SLD_CTA_IND_AHO_PEN_M_3_4), 
                            SUM(NO_AFILIADOS_ACTIVOS_M_3_4), 
                            SUM(NO_AFILIADOS_COTIZANTES_M_3_4),
                            SUM(NO_AFLDS_COBERT_SEG_PREV_M_3_4), 
                            SUM(TOTAL_AFILIADOS_H_4_8), 
                            SUM(PROM_SEMANAS_COTIZADAS_H_4_8),
                            SUM(PROM_SLD_CTA_IND_AHO_PEN_H_4_8), 
                            SUM(NO_AFILIADOS_ACTIVOS_H_4_8), 
                            SUM(NO_AFILIADOS_COTIZANTES_H_4_8),
                            SUM(NO_AFLDS_COBERT_SEG_PREV_H_4_8), 
                            SUM(TOTAL_AFILIADOS_M_4_8), 
                            SUM(PROM_SEMANAS_COTIZADAS_M_4_8),
                            SUM(PROM_SLD_CTA_IND_AHO_PEN_M_4_8), 
                            SUM(NO_AFILIADOS_ACTIVOS_M_4_8), 
                            SUM(NO_AFILIADOS_COTIZANTES_M_4_8),
                            SUM( NO_AFLDS_COBERT_SEG_PREV_M_4_8), 
                            SUM(TOTAL_AFILIADOS_H_8_12), 
                            SUM(PROM_SEMANAS_COTIZADAS_H_8_12),
                            SUM(PRM_SLD_CTA_IND_AHO_PEN_H_8_12), 
                            SUM(NO_AFILIADOS_ACTIVOS_H_8_12),
                            SUM(NO_AFILIADOS_COTIZANTES_H_8_12), 
                            SUM(NO_AFLDS_COBER_SEG_PREV_H_8_12),
                            SUM(TOTAL_AFILIADOS_M_8_12), 
                            SUM(PROM_SEMANAS_COTIZADAS_M_8_12), 
                            SUM(PRM_SLD_CTA_IND_AHO_PEN_M_8_12),
                            SUM(NO_AFILIADOS_ACTIVOS_M_8_12), 
                            SUM(NO_AFILIADOS_COTIZANTES_M_8_12),
                            SUM(NO_AFLDS_COBER_SEG_PREV_M_8_12), 
                            SUM(TOTAL_AFILIADOS_H_12_16), 
                            SUM(PROM_SEMANAS_COTIZADAS_H_12_16),
                            SUM(PRM_SLD_CTA_IND_AHOPEN_H_12_16), 
                            SUM(NO_AFILIADOS_ACTIVOS_H_12_16),
                            SUM(NO_AFLDS_COTIZANTES_H_12_16), 
                            SUM(NO_AFLDS_COBER_SEG_PRE_H_12_16),
                            SUM(TOTAL_AFILIADOS_M_12_16), 
                            SUM(PROM_SEMANAS_COTIZADAS_M_12_16), 
                            SUM(PRM_SLD_CTA_IND_AHOPEN_M_12_16),
                            SUM(NO_AFILIADOS_ACTIVOS_M_12_16), 
                            SUM(NO_AFLDS_COTIZANTES_M_12_16), 
                            SUM(NO_AFLDS_COBER_SEG_PRE_M_12_16),
                            SUM(TOTAL_AFILIADOS_H_16_20), 
                            SUM(PROM_SEMANAS_COTIZADAS_H_16_20), 
                            SUM(PRM_SLD_CTA_IND_AHOPEN_H_16_20),
                            SUM(NO_AFILIADOS_ACTIVOS_H_16_20), 
                            SUM(NO_AFLDS_COTIZANTES_H_16_20), 
                            SUM(NO_AFLDS_COBER_SEG_PRE_H_16_20),
                            SUM(TOTAL_AFILIADOS_M_16_20), 
                            SUM(PROM_SEMANAS_COTIZADAS_M_16_20), 
                            SUM(PRM_SLD_CTA_IND_AHOPEN_M_16_20),
                            SUM(NO_AFILIADOS_ACTIVOS_M_16_20),
                            SUM(NO_AFLDS_COTIZANTES_M_16_20), 
                            SUM(NO_AFLDS_COBER_SEG_PRE_M_16_20),
                            SUM(TOTAL_AFILIADOS_H_20), 
                            SUM(PROM_SEMANAS_COTIZADAS_H_20), 
                            SUM(PROM_SLD_CTA_IND_AHO_PEN_H_20),
                            SUM(NO_AFILIADOS_ACTIVOS_H_20), 
                            SUM(NO_AFILIADOS_COTIZANTES_H_20), 
                            SUM(NO_AFLDS_COBERT_SEG_PREV_H_20),
                            SUM(TOTAL_AFILIADOS_M_20), 
                            SUM(PROM_SEMANAS_COTIZADAS_M_20),
                            SUM(PROM_SLD_CTA_IND_AHO_PEN_M_20),
                            SUM(NO_AFILIADOS_ACTIVOS_M_20), 
                            SUM(NO_AFILIADOS_COTIZANTES_M_20), 
                            SUM(NO_AFLDS_COBERT_SEG_PREV_M_20),
                            SUM(TOTAL_AFILIADOS_H), 
                            SUM(TOTAL_AFILIADOS_M), 
                            SUM(TOTAL_AFILIADOS_TOTAL),
                            SUM(TOTAL_AFILIADOS_ACTIVOS_H), 
                            SUM(TOTAL_AFILIADOS_ACTIVOS_M), 
                            SUM(TOTAL_AFILIADOS_ACTIVOS_TOTAL),
                            SUM(TOTAL_AFILIADOS_COTIZANTES_H), 
                            SUM(TOTAL_AFILIADOS_COTIZANTES_M), 
                            SUM(TOTAL_AFILIADOS_COTIZANTES),
                            SUM(TIPO_AFILIACION_DEPENDIENTE), 
                            SUM(TIPO_AFILIACION_INDEPENDIENTE), 
                            SUM(ORIGEN_DE_AFILIACION_ISS),
                            SUM(ORIGEN_DE_AFILIACION_CAJAS), 
                            SUM(ORIGEN_AFILIACION_INGRESO), 
                            SUM(ORIGEN_AFILIACION_TRASLADO)
FROM                PROD_DWH_CONSULTA.FORMATO491
WHERE             UNIDAD_CAPTURA =3
                              ---FECBAL BETWEEN '2019-01-12' AND '2020-03-31'
                              ---FECBAL = '2020-01-31' 
GROUP BY      TIPO_ENTIDAD, 
                            CODIGO_ENTIDAD, 
                            NOMBRE_ENTIDAD, 
                            CIDT, 
                            FECBAL,
                            TIPO_IDENTIFICACION, 
                            NUMERO_IDENTIFICACION,
                             UNIDAD_CAPTURA, 
                            NOMBRE_UNIDAD_CAPTURA,
                            RENGLON
--ORDER BY       FECBAL,
 --                           CODIGO_ENTIDAD
--ORDER BY       FECBAL,
 --                           CODIGO_ENTIDAD,
  --                          SUBSTR ( NUMERO_IDENTIFICACION, 9 , 4 )
  UNION

--    TIPO DE FONDO MODERADO (CONVERGENCIA MODERADO Y MAYOR RIESGO) 
SELECT            TIPO_ENTIDAD, 
                            CODIGO_ENTIDAD, 
                            NOMBRE_ENTIDAD, 
                            CIDT, 
                            FECBAL,
                            TIPO_IDENTIFICACION, 
                            NUMERO_IDENTIFICACION,
                            'MOD Y MAY RIES' Tipo_de_Fondo,            
                            UNIDAD_CAPTURA, 
                            NOMBRE_UNIDAD_CAPTURA,
                            RENGLON, 
                            CASE WHEN RENGLON = 5 THEN '42 años'
                                        WHEN RENGLON = 10 THEN '43 años'
                                        WHEN RENGLON = 15 THEN '44 años'
                                        WHEN RENGLON = 20 THEN '45 años'
                                        WHEN RENGLON = 25 THEN '46 años'
                                        WHEN RENGLON = 30 THEN '47 años'
                                        WHEN RENGLON = 35 THEN '48 años'
                                        WHEN RENGLON = 40 THEN '49 años'
                                        WHEN RENGLON = 45 THEN '50 años'
                                        WHEN RENGLON = 50 THEN '51 o Más años'
                                        WHEN RENGLON = 999 THEN 'Total'
                            END Edades,
                            SUM(TOTAL_AFILIADOS_H_1), 
                            SUM(PROMEDIO_SEMANAS_COTIZADAS_H_1),
                            SUM(PROM_SLD_CTA_INDIV_AHO_PEN_H_1), 
                            SUM(NO_AFILIADOS_ACTIVOS_H_1), 
                            SUM(NO_AFILIADOS_COTIZANTES_H_1),
                            SUM(NO_AFLDS_COBERT_SEG_PREV_H_1), 
                            SUM(TOTAL_AFILIADOS_M_1), 
                            SUM(PROMEDIO_SEMANAS_COTIZADAS_M_1),
                            SUM(PROM_SLD_CTA_INDIV_AHO_PEN_M_1), 
                            SUM(NO_AFILIADOS_ACTIVOS_M_1), 
                            SUM(NO_AFILIADOS_COTIZANTES_M_1),
                            SUM(NO_AFLDS_COBERT_SEG_PREV_M_1), 
                            SUM(TOTAL_AFILIADOS_H_1_2), 
                            SUM( PROM_SEMANAS_COTIZADAS_H_1_2),
                            SUM(PROM_SLD_CTA_IND_AHO_PEN_H_1_2), 
                            SUM(NO_AFILIADOS_ACTIVOS_H_1_2), 
                            SUM(NO_AFILIADOS_COTIZANTES_H_1_2),
                            SUM(NO_AFLDS_COBERT_SEG_PREV_H_1_2), 
                            SUM(TOTAL_AFILIADOS_M_1_2), 
                            SUM(PROM_SEMANAS_COTIZADAS_M_1_2),
                            SUM(PROM_SLD_CTA_IND_AHO_PEN_M_1_2), 
                            SUM(NO_AFILIADOS_ACTIVOS_M_1_2), 
                            SUM(NO_AFILIADOS_COTIZANTES_M_1_2),
                            SUM(NO_AFLDS_COBERT_SEG_PREV_M_1_2), 
                            SUM(TOTAL_AFILIADOS_H_2_3), 
                            SUM(PROM_SEMANAS_COTIZADAS_H_2_3),
                            SUM(PROM_SLD_CTA_IND_AHO_PEN_H_2_3), 
                            SUM(NO_AFILIADOS_ACTIVOS_H_2_3), 
                            SUM(NO_AFILIADOS_COTIZANTES_H_2_3),
                            SUM(NO_AFLDS_COBERT_SEG_PREV_H_2_3), 
                            SUM(TOTAL_AFILIADOS_M_2_3), 
                            SUM(PROM_SEMANAS_COTIZADAS_M_2_3),
                            SUM(PROM_SLD_CTA_IND_AHO_PEN_M_2_3), 
                            SUM(NO_AFILIADOS_ACTIVOS_M_2_3), 
                            SUM(NO_AFILIADOS_COTIZANTES_M_2_3),
                            SUM(NO_AFLDS_COBERT_SEG_PREV_M_2_3), 
                            SUM(TOTAL_AFILIADOS_H_3_4), 
                            SUM(PROM_SEMANAS_COTIZADAS_H_3_4),
                            SUM(PROM_SLD_CTA_IND_AHO_PEN_H_3_4), 
                            SUM(NO_AFILIADOS_ACTIVOS_H_3_4), 
                            SUM(NO_AFILIADOS_COTIZANTES_H_3_4),
                            SUM(NO_AFLDS_COBERT_SEG_PREV_H_3_4), 
                            SUM(TOTAL_AFILIADOS_M_3_4), 
                            SUM(PROM_SEMANAS_COTIZADAS_M_3_4),
                            SUM(PROM_SLD_CTA_IND_AHO_PEN_M_3_4), 
                            SUM(NO_AFILIADOS_ACTIVOS_M_3_4), 
                            SUM(NO_AFILIADOS_COTIZANTES_M_3_4),
                            SUM(NO_AFLDS_COBERT_SEG_PREV_M_3_4), 
                            SUM(TOTAL_AFILIADOS_H_4_8), 
                            SUM(PROM_SEMANAS_COTIZADAS_H_4_8),
                            SUM(PROM_SLD_CTA_IND_AHO_PEN_H_4_8), 
                            SUM(NO_AFILIADOS_ACTIVOS_H_4_8), 
                            SUM(NO_AFILIADOS_COTIZANTES_H_4_8),
                            SUM(NO_AFLDS_COBERT_SEG_PREV_H_4_8), 
                            SUM(TOTAL_AFILIADOS_M_4_8), 
                            SUM(PROM_SEMANAS_COTIZADAS_M_4_8),
                            SUM(PROM_SLD_CTA_IND_AHO_PEN_M_4_8), 
                            SUM(NO_AFILIADOS_ACTIVOS_M_4_8), 
                            SUM(NO_AFILIADOS_COTIZANTES_M_4_8),
                            SUM( NO_AFLDS_COBERT_SEG_PREV_M_4_8), 
                            SUM(TOTAL_AFILIADOS_H_8_12), 
                            SUM(PROM_SEMANAS_COTIZADAS_H_8_12),
                            SUM(PRM_SLD_CTA_IND_AHO_PEN_H_8_12), 
                            SUM(NO_AFILIADOS_ACTIVOS_H_8_12),
                            SUM(NO_AFILIADOS_COTIZANTES_H_8_12), 
                            SUM(NO_AFLDS_COBER_SEG_PREV_H_8_12),
                            SUM(TOTAL_AFILIADOS_M_8_12), 
                            SUM(PROM_SEMANAS_COTIZADAS_M_8_12), 
                            SUM(PRM_SLD_CTA_IND_AHO_PEN_M_8_12),
                            SUM(NO_AFILIADOS_ACTIVOS_M_8_12), 
                            SUM(NO_AFILIADOS_COTIZANTES_M_8_12),
                            SUM(NO_AFLDS_COBER_SEG_PREV_M_8_12), 
                            SUM(TOTAL_AFILIADOS_H_12_16), 
                            SUM(PROM_SEMANAS_COTIZADAS_H_12_16),
                            SUM(PRM_SLD_CTA_IND_AHOPEN_H_12_16), 
                            SUM(NO_AFILIADOS_ACTIVOS_H_12_16),
                            SUM(NO_AFLDS_COTIZANTES_H_12_16), 
                            SUM(NO_AFLDS_COBER_SEG_PRE_H_12_16),
                            SUM(TOTAL_AFILIADOS_M_12_16), 
                            SUM(PROM_SEMANAS_COTIZADAS_M_12_16), 
                            SUM(PRM_SLD_CTA_IND_AHOPEN_M_12_16),
                            SUM(NO_AFILIADOS_ACTIVOS_M_12_16), 
                            SUM(NO_AFLDS_COTIZANTES_M_12_16), 
                            SUM(NO_AFLDS_COBER_SEG_PRE_M_12_16),
                            SUM(TOTAL_AFILIADOS_H_16_20), 
                            SUM(PROM_SEMANAS_COTIZADAS_H_16_20), 
                            SUM(PRM_SLD_CTA_IND_AHOPEN_H_16_20),
                            SUM(NO_AFILIADOS_ACTIVOS_H_16_20), 
                            SUM(NO_AFLDS_COTIZANTES_H_16_20), 
                            SUM(NO_AFLDS_COBER_SEG_PRE_H_16_20),
                            SUM(TOTAL_AFILIADOS_M_16_20), 
                            SUM(PROM_SEMANAS_COTIZADAS_M_16_20), 
                            SUM(PRM_SLD_CTA_IND_AHOPEN_M_16_20),
                            SUM(NO_AFILIADOS_ACTIVOS_M_16_20),
                            SUM(NO_AFLDS_COTIZANTES_M_16_20), 
                            SUM(NO_AFLDS_COBER_SEG_PRE_M_16_20),
                            SUM(TOTAL_AFILIADOS_H_20), 
                            SUM(PROM_SEMANAS_COTIZADAS_H_20), 
                            SUM(PROM_SLD_CTA_IND_AHO_PEN_H_20),
                            SUM(NO_AFILIADOS_ACTIVOS_H_20), 
                            SUM(NO_AFILIADOS_COTIZANTES_H_20), 
                            SUM(NO_AFLDS_COBERT_SEG_PREV_H_20),
                            SUM(TOTAL_AFILIADOS_M_20), 
                            SUM(PROM_SEMANAS_COTIZADAS_M_20),
                            SUM(PROM_SLD_CTA_IND_AHO_PEN_M_20),
                            SUM(NO_AFILIADOS_ACTIVOS_M_20), 
                            SUM(NO_AFILIADOS_COTIZANTES_M_20), 
                            SUM(NO_AFLDS_COBERT_SEG_PREV_M_20),
                            SUM(TOTAL_AFILIADOS_H), 
                            SUM(TOTAL_AFILIADOS_M), 
                            SUM(TOTAL_AFILIADOS_TOTAL),
                            SUM(TOTAL_AFILIADOS_ACTIVOS_H), 
                            SUM(TOTAL_AFILIADOS_ACTIVOS_M), 
                            SUM(TOTAL_AFILIADOS_ACTIVOS_TOTAL),
                            SUM(TOTAL_AFILIADOS_COTIZANTES_H), 
                            SUM(TOTAL_AFILIADOS_COTIZANTES_M), 
                            SUM(TOTAL_AFILIADOS_COTIZANTES),
                            SUM(TIPO_AFILIACION_DEPENDIENTE), 
                            SUM(TIPO_AFILIACION_INDEPENDIENTE), 
                            SUM(ORIGEN_DE_AFILIACION_ISS),
                            SUM(ORIGEN_DE_AFILIACION_CAJAS), 
                            SUM(ORIGEN_AFILIACION_INGRESO), 
                            SUM(ORIGEN_AFILIACION_TRASLADO)
FROM                PROD_DWH_CONSULTA.FORMATO491
WHERE             UNIDAD_CAPTURA =4 
                             ---- FECBAL BETWEEN '2019-01-12' AND '2020-03-31'
                             --FECBAL = '2020-01-31' 
GROUP BY      TIPO_ENTIDAD, 
                            CODIGO_ENTIDAD, 
                            NOMBRE_ENTIDAD, 
                            CIDT, 
                            FECBAL,
                            TIPO_IDENTIFICACION, 
                            NUMERO_IDENTIFICACION,
                             UNIDAD_CAPTURA, 
                            NOMBRE_UNIDAD_CAPTURA,
                            RENGLON
ORDER BY 5,7,9,11
--query493
SELECT    TIPO_ENTIDAD, 
                    CODIGO_ENTIDAD, 
                    NOMBRE_ENTIDAD, 
                    FECHA_CORTE,
                    TIPO_IDENTIFICACION,
                    IDENTIFICACION, 
                    CASE WHEN SUBSTR(IDENTIFICACION,9,4)  = 1000 THEN 'MODERADO'
                                WHEN SUBSTR(IDENTIFICACION,9,4)  = 5000 THEN 'CONSERVADOR'
                                WHEN SUBSTR(IDENTIFICACION,9,4)  =  6000 THEN 'MAYOR RIESGO'
                                WHEN SUBSTR(IDENTIFICACION,9,4)  =  8000 THEN 'OM_ALTERNATIVO'
                     END AS Tipo_Fondo,           
                    UNIDAD_CAPTURA, 
                    NOMBRE_UNIDAD_CAPTURA,
                    RENGLON, 
                    MUJERES_RANGO_EDAD_31, 
                    MUJERES_RANGO_EDAD_31_36, 
                    MUJERES_RANGO_EDAD_36_41,
                    MUJERES_RANGO_EDAD_41_46, 
                    MUJERES_RANGO_EDAD_46, 
                    HOMBRES_RANGO_EDAD_36,
                    HOMBRES_RANGO_EDAD_36_41, 
                    HOMBRES_RANGO_EDAD_41_46, 
                    HOMBRES_RANGO_EDAD_46_51,
                    HOMBRES_RANGO_EDAD_51, 
                    TOTAL_AFILIADOS, 
                    VALOR, 
                    TIPO_INFORME,
                    CIDT
FROM        PROD_DWH_CONSULTA.S9_FORMATO_493
WHERE     UNIDAD_CAPTURA=1
UNION
SELECT    TIPO_ENTIDAD, 
                    CODIGO_ENTIDAD, 
                    NOMBRE_ENTIDAD, 
                    FECHA_CORTE,
                    TIPO_IDENTIFICACION,
                    IDENTIFICACION, 
                    'CONVERGENCIA-CONSERVADOR Y MODERADO' Tipo_Fondo,
                    UNIDAD_CAPTURA, 
                    NOMBRE_UNIDAD_CAPTURA,
                    RENGLON, 
                    MUJERES_RANGO_EDAD_31, 
                    MUJERES_RANGO_EDAD_31_36, 
                    MUJERES_RANGO_EDAD_36_41,
                    MUJERES_RANGO_EDAD_41_46, 
                    MUJERES_RANGO_EDAD_46, 
                    HOMBRES_RANGO_EDAD_36,
                    HOMBRES_RANGO_EDAD_36_41, 
                    HOMBRES_RANGO_EDAD_41_46, 
                    HOMBRES_RANGO_EDAD_46_51,
                    HOMBRES_RANGO_EDAD_51, 
                    TOTAL_AFILIADOS, 
                    VALOR, 
                    TIPO_INFORME,
                    CIDT
FROM        PROD_DWH_CONSULTA.S9_FORMATO_493
WHERE     UNIDAD_CAPTURA=2
UNION
SELECT    TIPO_ENTIDAD, 
                    CODIGO_ENTIDAD, 
                    NOMBRE_ENTIDAD, 
                    FECHA_CORTE,
                    TIPO_IDENTIFICACION,
                    IDENTIFICACION, 
                    'CONVERGENCIA-CONSERVADOR Y MAYOR RIESGO' Tipo_Fondo,
                    UNIDAD_CAPTURA, 
                    NOMBRE_UNIDAD_CAPTURA,
                    RENGLON, 
                    MUJERES_RANGO_EDAD_31, 
                    MUJERES_RANGO_EDAD_31_36, 
                    MUJERES_RANGO_EDAD_36_41,
                    MUJERES_RANGO_EDAD_41_46, 
                    MUJERES_RANGO_EDAD_46, 
                    HOMBRES_RANGO_EDAD_36,
                    HOMBRES_RANGO_EDAD_36_41, 
                    HOMBRES_RANGO_EDAD_41_46, 
                    HOMBRES_RANGO_EDAD_46_51,
                    HOMBRES_RANGO_EDAD_51, 
                    TOTAL_AFILIADOS, 
                    VALOR, 
                    TIPO_INFORME,
                    CIDT
FROM        PROD_DWH_CONSULTA.S9_FORMATO_493
WHERE     UNIDAD_CAPTURA=3
UNION
SELECT    TIPO_ENTIDAD, 
                    CODIGO_ENTIDAD, 
                    NOMBRE_ENTIDAD, 
                    FECHA_CORTE,
                    TIPO_IDENTIFICACION,
                    IDENTIFICACION, 
                    'MODERADO(Traslados Recibidos)' Tipo_Fondo,
                    UNIDAD_CAPTURA, 
                    NOMBRE_UNIDAD_CAPTURA,
                    RENGLON, 
                    MUJERES_RANGO_EDAD_31, 
                    MUJERES_RANGO_EDAD_31_36, 
                    MUJERES_RANGO_EDAD_36_41,
                    MUJERES_RANGO_EDAD_41_46, 
                    MUJERES_RANGO_EDAD_46, 
                    HOMBRES_RANGO_EDAD_36,
                    HOMBRES_RANGO_EDAD_36_41, 
                    HOMBRES_RANGO_EDAD_41_46, 
                    HOMBRES_RANGO_EDAD_46_51,
                    HOMBRES_RANGO_EDAD_51, 
                    TOTAL_AFILIADOS, 
                    VALOR, 
                    TIPO_INFORME,
                    CIDT
FROM        PROD_DWH_CONSULTA.S9_FORMATO_493
WHERE     UNIDAD_CAPTURA=4
UNION
SELECT    TIPO_ENTIDAD, 
                    CODIGO_ENTIDAD, 
                    NOMBRE_ENTIDAD, 
                    FECHA_CORTE,
                    TIPO_IDENTIFICACION,
                    IDENTIFICACION, 
                    'MODERADO(Comisiones)' Tipo_Fondo,
                    UNIDAD_CAPTURA, 
                    NOMBRE_UNIDAD_CAPTURA,
                    RENGLON, 
                    MUJERES_RANGO_EDAD_31, 
                    MUJERES_RANGO_EDAD_31_36, 
                    MUJERES_RANGO_EDAD_36_41,
                    MUJERES_RANGO_EDAD_41_46, 
                    MUJERES_RANGO_EDAD_46, 
                    HOMBRES_RANGO_EDAD_36,
                    HOMBRES_RANGO_EDAD_36_41, 
                    HOMBRES_RANGO_EDAD_41_46, 
                    HOMBRES_RANGO_EDAD_46_51,
                    HOMBRES_RANGO_EDAD_51, 
                    TOTAL_AFILIADOS, 
                    VALOR, 
                    TIPO_INFORME,
                    CIDT
FROM        PROD_DWH_CONSULTA.S9_FORMATO_493
WHERE     UNIDAD_CAPTURA=5
UNION
SELECT    TIPO_ENTIDAD, 
                    CODIGO_ENTIDAD, 
                    NOMBRE_ENTIDAD, 
                    FECHA_CORTE,
                    TIPO_IDENTIFICACION,
                    IDENTIFICACION, 
                    'MODERADO(Comisiones)' Tipo_Fondo,
                    UNIDAD_CAPTURA, 
                    NOMBRE_UNIDAD_CAPTURA,
                    RENGLON, 
                    MUJERES_RANGO_EDAD_31, 
                    MUJERES_RANGO_EDAD_31_36, 
                    MUJERES_RANGO_EDAD_36_41,
                    MUJERES_RANGO_EDAD_41_46, 
                    MUJERES_RANGO_EDAD_46, 
                    HOMBRES_RANGO_EDAD_36,
                    HOMBRES_RANGO_EDAD_36_41, 
                    HOMBRES_RANGO_EDAD_41_46, 
                    HOMBRES_RANGO_EDAD_46_51,
                    HOMBRES_RANGO_EDAD_51, 
                    TOTAL_AFILIADOS, 
                    VALOR, 
                    TIPO_INFORME,
                    CIDT
FROM        PROD_DWH_CONSULTA.S9_FORMATO_493
WHERE     UNIDAD_CAPTURA=6
--query495
SELECT	TIPO_ENTIDAD, CODIGO_ENTIDAD, NOMBRE_ENTIDAD, FECHA_CORTE,
		TIPO_IDENTIFICACION, IDENTIFICACION, UNIDAD_CAPTURA, NOMBRE_UNIDAD_CAPTURA,
		RENGLON, RTR_PRGRMD_V_H, RTR_PRGRMD_V_M, RTR_PRGRMD_I_H, RTR_PRGRMD_I_M,
		RTR_PRGRMD_I, RTR_PRGRMD_S_H, RTR_PRGRMD_S_M, RTR_PRGRMD_S, RTR_PRGRMD_RNT_VTLC_DIF_V_H,
		RTR_PRGRMD_RNT_VTLC_DIF_V_M, RTR_PRGRMD_RNT_VTLC_DIF_I_H, RTR_PRGRMD_RNT_VTLC_DIF_I_M,
		RTR_PRGRMD_RNT_VTLC_DIF_I, RTR_PRGRMD_RNT_VTLC_DIF_S_H, RTR_PRGRMD_RNT_VTLC_DIF_S_M,
		RTR_PRGRMD_RNT_VTLC_DIF_S, PNSNS_FLL_JDCL_V_H, PNSNS_FLL_JDCL_V_M,
		PNSNS_FLL_JDCL_I_H, PNSNS_FLL_JDCL_I_M, PNSNS_FLL_JDCL_I, PNSNS_FLL_JDCL_S_H,
		PNSNS_FLL_JDCL_S_M, PNSNS_FLL_JDCL_S, RNT_TMP_CRT_VTLC_DFM_CRT_V_H,
		RNT_TMP_CRT_VTLC_DFM_CRT_V_M, RNT_TMP_CRT_VTLC_DFM_CRT_I_H, RNT_TMP_CRT_VTLC_DFM_CRT_I_M,
		RNT_TMP_CRT_RNT_VTLC_DFM_CRT_I, RNT_TMP_CRT_VTLC_DFM_CRT_S_H,
		RNT_TMP_CRT_VTLC_DFM_CRT_S_M, RNT_TMP_CRT_RNT_VTLC_DFM_CRT_S,
		RNT_VTLC_INMDT_V_H, RNT_VTLC_INMDT_V_M, RNT_VTLC_INMDT_I_H, RNT_VTLC_INMDT_I_M,
		RNT_VTLC_INMDT_I, RNT_VTLC_INMDT_S_H, RNT_VTLC_INMDT_S_M, RNT_VTLC_INMDT_S,
		RNT_TMP_VRBL_RNT_VTLC_DIF_V_H, RNT_TMP_VRBL_RNT_VTLC_DIF_V_M,
		RNT_TMP_VRBL_RNT_VTLC_DIF_I_H, RNT_TMP_VRBL_RNT_VTLC_DIF_I_M,
		RNT_TMP_VRBL_RNT_VTLC_DIF_I, RNT_TMP_VRBL_RNT_VTLC_DIF_S_H, RNT_TMP_VRBL_RNT_VTLC_DIF_S_M,
		RNT_TMP_VRBL_RNT_VTLC_DIF_S, RNT_TMP_VRBL_RNT_VTLC_INMD_V_H,
		RNT_TMP_VRBL_RNT_VTLC_INMD_V_M, RNT_TMP_VRBL_RNT_VTLC_INMD_I_H,
		RNT_TMP_VRBL_RNT_VTLC_INMD_I_M, RNT_TMP_VRBL_RNT_VTLC_INMDT_I,
		RNT_TMP_VRBL_RNT_VTLC_INMD_S_H, RNT_TMP_VRBL_RNT_VTLC_INMD_S_M,
		RNT_TMP_VRBL_RNT_VTLC_INMDT_S, RTR_PRGRMD_SIN_NGCCN_BONO_V_H,
		RTR_PRGRMD_SIN_NGCCN_BONO_V_M, TIPO_INFORME, CIDT
FROM PROD_DWH_CONSULTA.S9_FORMATO_495		 
