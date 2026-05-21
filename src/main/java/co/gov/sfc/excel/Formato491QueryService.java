package co.gov.sfc.excel;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.jdbc.core.JdbcTemplate;
import org.springframework.stereotype.Component;

import java.math.BigDecimal;
import java.sql.Date;
import java.time.LocalDate;

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

        BigDecimal menores30 = scalar("afiliados_menor30", sqlMenor30(), fecha);

        BigDecimal afiliados30a44 = scalar("afiliados_30_44", sql30a44(), fecha);

        BigDecimal afiliados45a59 = scalar("afiliados_45_59", sql45a59(), fecha);

        BigDecimal afiliadosMayor60 = scalar("afiliados_mayor_60", sqlMayor60(), fecha);

        return new Resumen491(afiliados, afiliadosActivos, menores30, afiliados30a44, afiliados45a59, afiliadosMayor60);
    }

    private BigDecimal scalar(String metric, String sql, Date fechaCorte) {
        log.info("Formato491QueryService ejecutando metric={} fechaCorte={} sql=\"{}\"",
                metric,
                fechaCorte,
                sql.replace("\n", " ").replaceAll("\\s+", " ").trim());
        BigDecimal value = jdbcTemplate.queryForObject(sql, BigDecimal.class, fechaCorte);
        log.info("Formato491QueryService resultado metric={} fechaCorte={} valor={}", metric, fechaCorte, value);
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

    public record Resumen491(
            BigDecimal afiliados,
            BigDecimal afiliadosActivos,
            BigDecimal afiliadosMenor30,
            BigDecimal afiliados30a44,
            BigDecimal afiliados45a59,
            BigDecimal afiliadosMayor60
    ) {}
}
