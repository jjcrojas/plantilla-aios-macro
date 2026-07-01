package co.gov.sfc.excel;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.jdbc.core.JdbcTemplate;
import org.springframework.stereotype.Component;

import java.math.BigDecimal;
import java.sql.Date;
import java.time.LocalDate;
import java.time.temporal.TemporalAdjusters;
import java.util.HashMap;
import java.util.Map;

@Component
public class Formato493QueryService {

    private static final Logger log = LoggerFactory.getLogger(Formato493QueryService.class);

    private final JdbcTemplate jdbcTemplate;

    public Formato493QueryService(JdbcTemplate jdbcTemplate) {
        this.jdbcTemplate = jdbcTemplate;
    }

    public BigDecimal leerTraspasosSistema(LocalDate fechaCorte) {
        Periodo periodo = periodoDoceMeses(fechaCorte);
        return scalar("traspasos_sistema", sqlTraspasos(null), Date.valueOf(periodo.inicio()), Date.valueOf(periodo.fin()));
    }

    public Map<String, BigDecimal> leerTraspasosPorEntidad(LocalDate fechaCorte) {
        Periodo periodo = periodoDoceMeses(fechaCorte);
        Map<String, BigDecimal> out = new HashMap<>();
        out.put("colf", leerTraspasosPorEntidad(periodo, 10, "colfondos"));
        out.put("porv", leerTraspasosPorEntidad(periodo, 3, "porvenir"));
        out.put("prot", leerTraspasosPorEntidad(periodo, 2, "proteccion"));
        out.put("sk", leerTraspasosPorEntidad(periodo, 9, "skandia"));
        log.info("Formato493QueryService resultado metric=traspasos_por_entidad fechaInicio={} fechaFin={} valores={}", periodo.inicio(), periodo.fin(), out);
        return out;
    }

    public BigDecimal leerFallecidosSistema(LocalDate fechaCorte) {
        Periodo periodo = periodoDoceMeses(fechaCorte);
        return scalar("fallecidos_sistema", sqlFallecidos(), Date.valueOf(periodo.inicio()), Date.valueOf(periodo.fin()));
    }

    private BigDecimal leerTraspasosPorEntidad(Periodo periodo, int codigoEntidad, String entidad) {
        return scalar("traspasos_" + entidad + "_codigo_entidad_" + codigoEntidad,
                sqlTraspasos("CODIGO_ENTIDAD = ?"),
                Date.valueOf(periodo.inicio()), Date.valueOf(periodo.fin()), codigoEntidad);
    }

    private Periodo periodoDoceMeses(LocalDate fechaCorte) {
        LocalDate fin = fechaCorte.with(TemporalAdjusters.lastDayOfMonth());
        LocalDate inicio = fin.minusMonths(11).with(TemporalAdjusters.lastDayOfMonth());
        return new Periodo(inicio, fin);
    }

    private BigDecimal scalar(String metric, String sql, Object... params) {
        log.info("Formato493QueryService ejecutando metric={} params={} sql=\"{}\"",
                metric,
                java.util.Arrays.toString(params),
                sql.replace("\n", " ").replaceAll("\\s+", " ").trim());
        BigDecimal value = jdbcTemplate.queryForObject(sql, BigDecimal.class, params);
        log.info("Formato493QueryService resultado metric={} valor={}", metric, value);
        return value == null ? BigDecimal.ZERO : value;
    }

    private String sqlTraspasos(String filtroEntidad) {
        String filtro = filtroEntidad == null ? "" : "\n  AND " + filtroEntidad;
        return ("""
                SELECT COALESCE(SUM(
                    COALESCE(MUJERES_RANGO_EDAD_31,0) +
                    COALESCE(MUJERES_RANGO_EDAD_31_36,0) +
                    COALESCE(MUJERES_RANGO_EDAD_36_41,0) +
                    COALESCE(MUJERES_RANGO_EDAD_41_46,0) +
                    COALESCE(MUJERES_RANGO_EDAD_46,0) +
                    COALESCE(HOMBRES_RANGO_EDAD_36,0) +
                    COALESCE(HOMBRES_RANGO_EDAD_36_41,0) +
                    COALESCE(HOMBRES_RANGO_EDAD_41_46,0) +
                    COALESCE(HOMBRES_RANGO_EDAD_46_51,0) +
                    COALESCE(HOMBRES_RANGO_EDAD_51,0)
                ), 0) AS TOTAL_PERSONAS
                FROM PROD_DWH_CONSULTA.S9_FORMATO_493
                WHERE FECHA_CORTE BETWEEN ? AND ?
                  AND (
                       (UNIDAD_CAPTURA = 1 AND RENGLON IN (70, 75, 90, 95))
                    OR (UNIDAD_CAPTURA = 2 AND RENGLON IN (40, 45, 60, 65))
                    OR (UNIDAD_CAPTURA = 3 AND RENGLON IN (40, 45, 60, 65))
                    OR (UNIDAD_CAPTURA = 6 AND RENGLON IN (35, 40, 45, 50))
                  )%s
                """).formatted(filtro);
    }

    private String sqlFallecidos() {
        return """
                SELECT COALESCE(SUM(
                    COALESCE(MUJERES_RANGO_EDAD_31,0) +
                    COALESCE(MUJERES_RANGO_EDAD_31_36,0) +
                    COALESCE(MUJERES_RANGO_EDAD_36_41,0) +
                    COALESCE(MUJERES_RANGO_EDAD_41_46,0) +
                    COALESCE(MUJERES_RANGO_EDAD_46,0) +
                    COALESCE(HOMBRES_RANGO_EDAD_36,0) +
                    COALESCE(HOMBRES_RANGO_EDAD_36_41,0) +
                    COALESCE(HOMBRES_RANGO_EDAD_41_46,0) +
                    COALESCE(HOMBRES_RANGO_EDAD_46_51,0) +
                    COALESCE(HOMBRES_RANGO_EDAD_51,0)
                ), 0) AS TOTAL_PERSONAS
                FROM PROD_DWH_CONSULTA.S9_FORMATO_493
                WHERE UNIDAD_CAPTURA = 1
                  AND RENGLON IN (165, 170, 175)
                  AND FECHA_CORTE BETWEEN ? AND ?
                """;
    }

    private record Periodo(LocalDate inicio, LocalDate fin) {}
}
