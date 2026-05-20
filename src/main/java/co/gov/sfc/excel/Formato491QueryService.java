package co.gov.sfc.excel;

import org.springframework.jdbc.core.JdbcTemplate;
import org.springframework.stereotype.Component;

import java.math.BigDecimal;
import java.sql.Date;
import java.time.LocalDate;

@Component
public class Formato491QueryService {

    private static final String FONDOS_FILTRO = "('1000','5000','6000','7000','8000')";

    private final JdbcTemplate jdbcTemplate;

    public Formato491QueryService(JdbcTemplate jdbcTemplate) {
        this.jdbcTemplate = jdbcTemplate;
    }

    public Resumen491 leerResumen(LocalDate fechaCorte) {
        Date fecha = Date.valueOf(fechaCorte);
        BigDecimal afiliados = scalar("""
                SELECT COALESCE(SUM(COALESCE(TOTAL_AFILIADOS_TOTAL, 0)), 0)
                FROM PROD_DWH_CONSULTA.FORMATO491
                WHERE FECBAL = ?
                  AND RENGLON = '999'
                  AND SUBSTR(NUMERO_IDENTIFICACION, 9, 4) IN %s
                """.formatted(FONDOS_FILTRO), fecha);

        BigDecimal afiliadosActivos = scalar("""
                SELECT COALESCE(SUM(COALESCE(TOTAL_AFILIADOS_ACTIVOS_TOTAL, 0)), 0)
                FROM PROD_DWH_CONSULTA.FORMATO491
                WHERE FECBAL = ?
                  AND RENGLON = '999'
                  AND SUBSTR(NUMERO_IDENTIFICACION, 9, 4) IN %s
                """.formatted(FONDOS_FILTRO), fecha);

        BigDecimal menores30 = scalar("""
                SELECT COALESCE(SUM(COALESCE(TOTAL_AFILIADOS_TOTAL, 0)), 0)
                FROM PROD_DWH_CONSULTA.FORMATO491
                WHERE FECBAL = ?
                  AND SUBSTR(NUMERO_IDENTIFICACION, 9, 4) IN %s
                  AND CAST(TRIM(UNIDAD_CAPTURA) AS INTEGER) = 1
                  AND CAST(TRIM(RENGLON) AS INTEGER) < 80
                """.formatted(FONDOS_FILTRO), fecha);

        BigDecimal afiliados30a44 = scalar("""
                SELECT COALESCE(SUM(COALESCE(TOTAL_AFILIADOS_TOTAL, 0)), 0)
                FROM PROD_DWH_CONSULTA.FORMATO491
                WHERE FECBAL = ?
                  AND SUBSTR(NUMERO_IDENTIFICACION, 9, 4) IN %s
                  AND (
                      (CAST(TRIM(UNIDAD_CAPTURA) AS INTEGER) = 1 AND CAST(TRIM(RENGLON) AS INTEGER) BETWEEN 80 AND 150)
                      OR
                      (CAST(TRIM(UNIDAD_CAPTURA) AS INTEGER) = 4 AND CAST(TRIM(RENGLON) AS INTEGER) BETWEEN 5 AND 15)
                  )
                """.formatted(FONDOS_FILTRO), fecha);

        BigDecimal afiliados45a59 = scalar("""
                SELECT COALESCE(SUM(COALESCE(TOTAL_AFILIADOS_TOTAL, 0)), 0)
                FROM PROD_DWH_CONSULTA.FORMATO491
                WHERE FECBAL = ?
                  AND SUBSTR(NUMERO_IDENTIFICACION, 9, 4) IN %s
                  AND (
                      (CAST(TRIM(RENGLON) AS INTEGER) BETWEEN 155 AND 225)
                      OR
                      (CAST(TRIM(UNIDAD_CAPTURA) AS INTEGER) > 1 AND CAST(TRIM(RENGLON) AS INTEGER) BETWEEN 20 AND 50)
                      OR
                      (CAST(TRIM(UNIDAD_CAPTURA) AS INTEGER) > 1 AND CAST(TRIM(UNIDAD_CAPTURA) AS INTEGER) < 4 AND CAST(TRIM(RENGLON) AS INTEGER) < 20)
                  )
                """.formatted(FONDOS_FILTRO), fecha);

        BigDecimal afiliadosMayor60 = scalar("""
                SELECT COALESCE(SUM(COALESCE(TOTAL_AFILIADOS_TOTAL, 0)), 0)
                FROM PROD_DWH_CONSULTA.FORMATO491
                WHERE FECBAL = ?
                  AND SUBSTR(NUMERO_IDENTIFICACION, 9, 4) IN %s
                  AND (
                      (CAST(TRIM(RENGLON) AS INTEGER) >= 230 AND CAST(TRIM(RENGLON) AS INTEGER) < 999)
                      OR
                      (CAST(TRIM(UNIDAD_CAPTURA) AS INTEGER) > 1 AND CAST(TRIM(RENGLON) AS INTEGER) BETWEEN 55 AND 80)
                  )
                """.formatted(FONDOS_FILTRO), fecha);

        return new Resumen491(afiliados, afiliadosActivos, menores30, afiliados30a44, afiliados45a59, afiliadosMayor60);
    }

    private BigDecimal scalar(String sql, Date fechaCorte) {
        BigDecimal value = jdbcTemplate.queryForObject(sql, BigDecimal.class, fechaCorte);
        return value == null ? BigDecimal.ZERO : value;
    }

    public record Resumen491(
            BigDecimal afiliados,
            BigDecimal afiliadosActivos,
            BigDecimal afiliadosMenor30,
            BigDecimal afiliados30a44,
            BigDecimal afiliados45a59,
            BigDecimal afiliadosMayor60
    ) {}
}
