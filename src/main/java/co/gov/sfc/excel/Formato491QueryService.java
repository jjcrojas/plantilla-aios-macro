package co.gov.sfc.excel;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.jdbc.core.JdbcTemplate;
import org.springframework.stereotype.Component;

import java.io.BufferedReader;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.math.BigDecimal;
import java.nio.charset.StandardCharsets;
import java.sql.Date;
import java.time.LocalDate;
import java.util.Comparator;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

@Component
public class Formato491QueryService {

    private static final Logger log = LoggerFactory.getLogger(Formato491QueryService.class);
    private static final String FONDOS_FILTRO = "('1000','5000','6000','7000','8000')";

    private final JdbcTemplate jdbcTemplate;

    public Formato491QueryService(JdbcTemplate jdbcTemplate) {
        this.jdbcTemplate = jdbcTemplate;
    }

    public Resumen491 leerResumen(LocalDate fechaCorte) {
        Date fecha = Date.valueOf(fechaCorte);
        BigDecimal afiliados = scalar("afiliados_total", sqlAfiliadosTotal(), fecha);

        BigDecimal afiliadosActivos = scalar("afiliados_activos_total", sqlAfiliadosActivosTotal(), fecha);

        BigDecimal mujeresAfiliadas = scalar("mujeres_afiliadas_total", sqlMujeresAfiliadasTotal(), fecha);

        BigDecimal menores30 = scalar("afiliados_menor30", sqlMenor30(), fecha);

        BigDecimal afiliados30a44 = scalar("afiliados_30_44", sql30a44(), fecha);

        BigDecimal afiliados45a59 = scalar("afiliados_45_59", sql45a59(), fecha);

        BigDecimal afiliadosMayor60 = scalar("afiliados_mayor_60", sqlMayor60(), fecha);

        BigDecimal aportantes = leerAportantesTotal(fechaCorte);
        BigDecimal aportantesSemestral = aportantes;
        BigDecimal concentracionAfiliados = leerConcentracionAfiliados(fechaCorte);
        BigDecimal salarioMinimoPonderadoCop = leerSalarioMinimoPonderadoCop(fechaCorte);

        return new Resumen491(afiliados, afiliadosActivos, mujeresAfiliadas, menores30, afiliados30a44, afiliados45a59, afiliadosMayor60, aportantes, aportantesSemestral, concentracionAfiliados, salarioMinimoPonderadoCop);
    }


    public BigDecimal leerAportantesTotal(LocalDate fechaCorte) {
        return scalar("aportantes_total", sqlAportantesTotal(), Date.valueOf(fechaCorte));
    }

    public BigDecimal leerAportantesSemestral(LocalDate fechaCorte) {
        return leerAportantesTotal(fechaCorte);
    }

    public BigDecimal leerSalarioMinimoPonderadoCop(LocalDate fechaCorte) {
        BigDecimal salarioMinimo = leerSalarioMinimoOficial(fechaCorte.getYear());
        Date fecha = Date.valueOf(fechaCorte);
        return scalar("salario_minimo_ponderado_cop", sqlSalarioMinimoPonderado(), fecha, fecha, salarioMinimo);
    }


    public BigDecimal leerConcentracionAfiliados(LocalDate fechaCorte) {
        Date fecha = Date.valueOf(fechaCorte);
        BigDecimal totalAfiliados = scalar("concentracion_afiliados_total_sistema", sqlAfiliadosTotal(), fecha);
        String sql = sqlAfiliadosPorEntidad();
        log.info("Formato491QueryService ejecutando metric=concentracion_afiliados_por_entidad params={} sql=\"{}\"",
                java.util.Arrays.toString(new Object[]{fecha}),
                sql.replace("\n", " ").replaceAll("\\s+", " ").trim());
        List<EntidadAfiliados> afiliadosPorEntidad = jdbcTemplate.query(sql, (rs, rowNum) -> new EntidadAfiliados(
                rs.getInt("CODIGO_ENTIDAD"),
                rs.getBigDecimal("AFILIADOS") == null ? BigDecimal.ZERO : rs.getBigDecimal("AFILIADOS")
        ), fecha);

        BigDecimal topDos = afiliadosPorEntidad.stream()
                .sorted(Comparator.comparing(EntidadAfiliados::afiliados).reversed())
                .limit(2)
                .map(EntidadAfiliados::afiliados)
                .reduce(BigDecimal.ZERO, BigDecimal::add);
        BigDecimal resultado = totalAfiliados.signum() == 0
                ? BigDecimal.ZERO
                : topDos.divide(totalAfiliados, 8, java.math.RoundingMode.HALF_UP).multiply(BigDecimal.valueOf(100));
        log.info("Formato491QueryService resultado metric=concentracion_afiliados fechaCorte={} topDos={} totalAfiliados={} valor={} entidades={}",
                fecha, topDos, totalAfiliados, resultado, afiliadosPorEntidad);
        return resultado;
    }

    public Map<String, BigDecimal> leerAportantesPorEntidad(LocalDate fechaCorte) {
        Date fecha = Date.valueOf(fechaCorte);
        Map<String, BigDecimal> out = new HashMap<>();
        out.put("colf", scalar("aportantes_trimestral_colfondos_codigo_entidad_10", sqlAportantesPorEntidad(), fecha, 10));
        out.put("porv", scalar("aportantes_trimestral_porvenir_codigo_entidad_3", sqlAportantesPorEntidad(), fecha, 3));
        out.put("prot", scalar("aportantes_trimestral_proteccion_codigo_entidad_2", sqlAportantesPorEntidad(), fecha, 2));
        out.put("sk", scalar("aportantes_trimestral_skandia_codigo_entidad_9", sqlAportantesPorEntidad(), fecha, 9));
        log.info("Formato491QueryService resultado metric=aportantes_trimestral_por_entidad fechaCorte={} valores={}", fecha, out);
        return out;
    }

    private BigDecimal scalar(String metric, String sql, Object... params) {
        log.info("Formato491QueryService ejecutando metric={} params={} sql=\"{}\"",
                metric,
                java.util.Arrays.toString(params),
                sql.replace("\n", " ").replaceAll("\\s+", " ").trim());
        BigDecimal value = jdbcTemplate.queryForObject(sql, BigDecimal.class, params);
        log.info("Formato491QueryService resultado metric={} valor={}", metric, value);
        return value == null ? BigDecimal.ZERO : value;
    }



    private String sqlAfiliadosTotal() {
        return """
                SELECT COALESCE(SUM(COALESCE(TOTAL_AFILIADOS_TOTAL, 0)), 0)
                FROM PROD_DWH_CONSULTA.FORMATO491
                WHERE FECBAL = ?
                  AND RENGLON = '999'
                  AND SUBSTR(NUMERO_IDENTIFICACION, 9, 4) IN %s
                """.formatted(FONDOS_FILTRO);
    }

    private String sqlAfiliadosActivosTotal() {
        return """
                SELECT COALESCE(SUM(COALESCE(TOTAL_AFILIADOS_ACTIVOS_TOTAL, 0)), 0)
                FROM PROD_DWH_CONSULTA.FORMATO491
                WHERE FECBAL = ?
                  AND RENGLON = '999'
                  AND SUBSTR(NUMERO_IDENTIFICACION, 9, 4) IN %s
                """.formatted(FONDOS_FILTRO);
    }

    private String sqlMujeresAfiliadasTotal() {
        return """
                SELECT COALESCE(SUM(COALESCE(TOTAL_AFILIADOS_M, 0)), 0)
                FROM PROD_DWH_CONSULTA.FORMATO491
                WHERE FECBAL = ?
                  AND RENGLON = '999'
                  AND SUBSTR(NUMERO_IDENTIFICACION, 9, 4) IN %s
                """.formatted(FONDOS_FILTRO);
    }

    private String sqlMenor30() {
        return """
                SELECT COALESCE(SUM(COALESCE(TOTAL_AFILIADOS_TOTAL, 0)), 0)
                FROM PROD_DWH_CONSULTA.FORMATO491
                WHERE FECBAL = ?
                  AND SUBSTR(NUMERO_IDENTIFICACION, 9, 4) IN %s
                  AND CAST(TRIM(UNIDAD_CAPTURA) AS INTEGER) = 1
                  AND CAST(TRIM(RENGLON) AS INTEGER) < 80
                """.formatted(FONDOS_FILTRO);
    }

    private String sql30a44() { return """
                SELECT COALESCE(SUM(COALESCE(TOTAL_AFILIADOS_TOTAL, 0)), 0)
                FROM PROD_DWH_CONSULTA.FORMATO491
                WHERE FECBAL = ?
                  AND SUBSTR(NUMERO_IDENTIFICACION, 9, 4) IN %s
                  AND ((CAST(TRIM(UNIDAD_CAPTURA) AS INTEGER) = 1 AND CAST(TRIM(RENGLON) AS INTEGER) BETWEEN 80 AND 150)
                      OR (CAST(TRIM(UNIDAD_CAPTURA) AS INTEGER) = 4 AND CAST(TRIM(RENGLON) AS INTEGER) BETWEEN 5 AND 15))
                """.formatted(FONDOS_FILTRO); }

    private String sql45a59() { return """
                SELECT COALESCE(SUM(COALESCE(TOTAL_AFILIADOS_TOTAL, 0)), 0)
                FROM PROD_DWH_CONSULTA.FORMATO491
                WHERE FECBAL = ?
                  AND SUBSTR(NUMERO_IDENTIFICACION, 9, 4) IN %s
                  AND ((CAST(TRIM(RENGLON) AS INTEGER) BETWEEN 155 AND 225)
                      OR (CAST(TRIM(UNIDAD_CAPTURA) AS INTEGER) > 1 AND CAST(TRIM(RENGLON) AS INTEGER) BETWEEN 20 AND 50)
                      OR (CAST(TRIM(UNIDAD_CAPTURA) AS INTEGER) > 1 AND CAST(TRIM(UNIDAD_CAPTURA) AS INTEGER) < 4 AND CAST(TRIM(RENGLON) AS INTEGER) < 20))
                """.formatted(FONDOS_FILTRO); }

    private String sqlMayor60() { return """
                SELECT COALESCE(SUM(COALESCE(TOTAL_AFILIADOS_TOTAL, 0)), 0)
                FROM PROD_DWH_CONSULTA.FORMATO491
                WHERE FECBAL = ?
                  AND SUBSTR(NUMERO_IDENTIFICACION, 9, 4) IN %s
                  AND ((CAST(TRIM(RENGLON) AS INTEGER) >= 230 AND CAST(TRIM(RENGLON) AS INTEGER) < 999)
                      OR (CAST(TRIM(UNIDAD_CAPTURA) AS INTEGER) > 1 AND CAST(TRIM(RENGLON) AS INTEGER) BETWEEN 55 AND 80))
                """.formatted(FONDOS_FILTRO); }



    private String sqlAportantesTotal() {
        return """
                SELECT COALESCE(SUM(COALESCE(TOTAL_AFILIADOS_COTIZANTES, 0)), 0)
                FROM PROD_DWH_CONSULTA.FORMATO491
                WHERE FECBAL = ?
                  AND RENGLON = '999'
                  AND SUBSTR(NUMERO_IDENTIFICACION, 9, 4) IN %s
                """.formatted(FONDOS_FILTRO);
    }

    private String sqlAportantesPorEntidad() {
        return """
                SELECT COALESCE(SUM(COALESCE(TOTAL_AFILIADOS_COTIZANTES, 0)), 0)
                FROM PROD_DWH_CONSULTA.FORMATO491
                WHERE FECBAL = ?
                  AND RENGLON = '999'
                  AND CODIGO_ENTIDAD = ?
                  AND SUBSTR(NUMERO_IDENTIFICACION, 9, 4) IN %s
                """.formatted(FONDOS_FILTRO);
    }



    private String sqlAfiliadosPorEntidad() {
        return """
                SELECT CODIGO_ENTIDAD,
                       COALESCE(SUM(COALESCE(TOTAL_AFILIADOS_TOTAL, 0)), 0) AS AFILIADOS
                FROM PROD_DWH_CONSULTA.FORMATO491
                WHERE FECBAL = ?
                  AND RENGLON = '999'
                  AND SUBSTR(NUMERO_IDENTIFICACION, 9, 4) IN %s
                GROUP BY CODIGO_ENTIDAD
                """.formatted(FONDOS_FILTRO);
    }

    private String sqlSalarioMinimoPonderado() {
        return """
                WITH base AS (
                    SELECT
                        COALESCE(SUM(
                            (COALESCE(TOTAL_AFILIADOS_H_1, 0) + COALESCE(TOTAL_AFILIADOS_M_1, 0)) * 1 +
                            (COALESCE(TOTAL_AFILIADOS_H_1_2, 0) + COALESCE(TOTAL_AFILIADOS_M_1_2, 0)) * 2 +
                            (COALESCE(TOTAL_AFILIADOS_H_2_3, 0) + COALESCE(TOTAL_AFILIADOS_M_2_3, 0)) * 3 +
                            (COALESCE(TOTAL_AFILIADOS_H_3_4, 0) + COALESCE(TOTAL_AFILIADOS_M_3_4, 0)) * 4 +
                            (COALESCE(TOTAL_AFILIADOS_H_4_8, 0) + COALESCE(TOTAL_AFILIADOS_M_4_8, 0)) * 8 +
                            (COALESCE(TOTAL_AFILIADOS_H_8_12, 0) + COALESCE(TOTAL_AFILIADOS_M_8_12, 0)) * 12 +
                            (COALESCE(TOTAL_AFILIADOS_H_12_16, 0) + COALESCE(TOTAL_AFILIADOS_M_12_16, 0)) * 16 +
                            (COALESCE(TOTAL_AFILIADOS_H_16_20, 0) + COALESCE(TOTAL_AFILIADOS_M_16_20, 0)) * 20 +
                            (COALESCE(TOTAL_AFILIADOS_H_20, 0) + COALESCE(TOTAL_AFILIADOS_M_20, 0)) * 25
                        ), 0) AS RANGOS_PONDERADOS
                    FROM PROD_DWH_CONSULTA.FORMATO491
                    WHERE FECBAL = ?
                      AND RENGLON = '999'
                      AND (
                          (CAST(TRIM(UNIDAD_CAPTURA) AS INTEGER) IN (1, 4) AND SUBSTR(NUMERO_IDENTIFICACION, 9, 4) = '1000') OR
                          (CAST(TRIM(UNIDAD_CAPTURA) AS INTEGER) IN (1, 2, 3) AND SUBSTR(NUMERO_IDENTIFICACION, 9, 4) = '5000') OR
                          (CAST(TRIM(UNIDAD_CAPTURA) AS INTEGER) IN (1) AND SUBSTR(NUMERO_IDENTIFICACION, 9, 4) = '6000')
                      )
                ), total_sistema AS (
                    SELECT COALESCE(SUM(COALESCE(TOTAL_AFILIADOS_TOTAL, 0)), 0) AS TOTAL_AFILIADOS
                    FROM PROD_DWH_CONSULTA.FORMATO491
                    WHERE FECBAL = ?
                      AND RENGLON = '999'
                      AND SUBSTR(NUMERO_IDENTIFICACION, 9, 4) IN %s
                )
                SELECT CASE
                         WHEN total_sistema.TOTAL_AFILIADOS = 0 THEN 0
                         ELSE (? * base.RANGOS_PONDERADOS) / total_sistema.TOTAL_AFILIADOS
                       END
                FROM base CROSS JOIN total_sistema
                """.formatted(FONDOS_FILTRO);
    }

    private BigDecimal leerSalarioMinimoOficial(int year) {
        try (InputStream in = salarioMinimoInputStream();
             BufferedReader reader = new BufferedReader(new InputStreamReader(in, StandardCharsets.UTF_8))) {
            String line;
            boolean header = true;
            while ((line = reader.readLine()) != null) {
                if (header) {
                    header = false;
                    continue;
                }
                if (line.isBlank()) {
                    continue;
                }
                String[] parts = line.trim().split("[\t,;]+");
                if (parts.length >= 2 && Integer.parseInt(parts[0].trim()) == year) {
                    return new BigDecimal(parts[1].trim());
                }
            }
        } catch (IOException e) {
            throw new IllegalStateException("No fue posible leer SalarioMinimo.csv", e);
        }
        throw new IllegalStateException("No existe salario mínimo configurado para el año " + year + " en SalarioMinimo.csv");
    }

    private InputStream salarioMinimoInputStream() throws IOException {
        InputStream classpath = getClass().getResourceAsStream("/SalarioMinimo.csv");
        if (classpath != null) {
            return classpath;
        }
        java.nio.file.Path local = java.nio.file.Path.of("SalarioMinimo.csv");
        if (java.nio.file.Files.exists(local)) {
            return java.nio.file.Files.newInputStream(local);
        }
        throw new IOException("No se encontró SalarioMinimo.csv en classpath ni en el directorio de trabajo");
    }

    private record EntidadAfiliados(Integer codigoEntidad, BigDecimal afiliados) {}

    public record Resumen491(
            BigDecimal afiliados,
            BigDecimal afiliadosActivos,
            BigDecimal mujeresAfiliadas,
            BigDecimal afiliadosMenor30,
            BigDecimal afiliados30a44,
            BigDecimal afiliados45a59,
            BigDecimal afiliadosMayor60,
            BigDecimal aportantes,
            BigDecimal aportantesSemestral,
            BigDecimal concentracionAfiliados,
            BigDecimal salarioMinimoPonderadoCop
    ) {}
}
