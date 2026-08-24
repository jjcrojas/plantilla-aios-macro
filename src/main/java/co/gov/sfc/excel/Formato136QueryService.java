package co.gov.sfc.excel;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.jdbc.core.JdbcTemplate;
import org.springframework.stereotype.Component;

import java.math.BigDecimal;
import java.sql.Date;
import java.time.LocalDate;
import java.util.HashMap;
import java.util.Locale;
import java.util.Map;

@Component
public class Formato136QueryService {

    private static final Logger log = LoggerFactory.getLogger(Formato136QueryService.class);
    private final JdbcTemplate jdbcTemplate;

    public Formato136QueryService(JdbcTemplate jdbcTemplate) {
        this.jdbcTemplate = jdbcTemplate;
    }

    public BigDecimal leerAportesRecibidos(LocalDate fechaCorte) {
        LocalDate fechaInicio = fechaCorte.minusYears(1).withDayOfMonth(1);
        return scalar("aportes_recibidos", sqlAportesRecibidos(), Date.valueOf(fechaInicio), Date.valueOf(fechaCorte));
    }

    public Map<String, BigDecimal> leerColombiaPorFondoEntidad(LocalDate fechaCorte) {
        Map<String, BigDecimal> out = new HashMap<>();
        leerFondoEstandar(out, "con", 5000, fechaCorte);
        leerFondoModerado(out, fechaCorte);
        leerFondoEstandar(out, "rp", 7000, fechaCorte);
        leerFondoEstandar(out, "mr", 6000, fechaCorte);
        log.info("Formato136QueryService resultado metric=colombia_por_fondo_entidad fechaCorte={} valores={}", fechaCorte, out);
        return out;
    }

    private void leerFondoEstandar(Map<String, BigDecimal> out, String fondo, int codigoPatrimonio, LocalDate fechaCorte) {
        String sql = sqlFondoEstandar();
        logQueryColombia(fondo, fechaCorte, sql, codigoPatrimonio);
        jdbcTemplate.query(sql, rs -> {
            String entidad = claveEntidad(rs.getInt("codigo_entidad"));
            if (entidad != null) {
                out.merge(fondo + "_" + entidad, nvl(rs.getBigDecimal("valor_mm_cop")), BigDecimal::add);
            }
        }, codigoPatrimonio, Date.valueOf(fechaCorte));
    }

    private void leerFondoModerado(Map<String, BigDecimal> out, LocalDate fechaCorte) {
        String sql = sqlFondoModerado();
        logQueryColombia("mod", fechaCorte, sql, 1000, 8000);
        jdbcTemplate.query(sql, rs -> {
            String claveReporte = rs.getString("clave_reporte");
            BigDecimal valor = nvl(rs.getBigDecimal("valor_mm_cop"));
            if ("SKANDIA_ALT".equals(normalize(claveReporte))) {
                out.merge("mod_alt", valor, BigDecimal::add);
                return;
            }
            String entidad = claveEntidad(rs.getInt("codigo_entidad"));
            if (entidad != null) {
                out.merge("mod_" + entidad, valor, BigDecimal::add);
            }
        }, Date.valueOf(fechaCorte));
    }

    private void logQueryColombia(String fondo, LocalDate fechaCorte, String sql, Object... reglasPatrimonio) {
        log.info("Formato136QueryService ejecutando metric=colombia_por_fondo fondo={} fechaCorte={} patrimonios={} sql=\"{}\"",
                fondo,
                fechaCorte,
                java.util.Arrays.toString(reglasPatrimonio),
                sql.replace("\n", " ").replaceAll("\\s+", " ").trim());
    }

    private BigDecimal nvl(BigDecimal value) {
        return value == null ? BigDecimal.ZERO : value;
    }

    private String claveEntidad(int codigoEntidad) {
        return switch (codigoEntidad) {
            case 10 -> "colf";
            case 3 -> "porv";
            case 2 -> "prot";
            case 9 -> "sk";
            default -> null;
        };
    }

    private String normalize(String value) {
        return value == null ? "" : value.trim().toUpperCase(Locale.ROOT);
    }

    private BigDecimal scalar(String metric, String sql, Object... params) {
        log.info("Formato136QueryService ejecutando metric={} params={} sql=\"{}\"",
                metric,
                java.util.Arrays.toString(params),
                sql.replace("\n", " ").replaceAll("\\s+", " ").trim());
        BigDecimal value = jdbcTemplate.queryForObject(sql, BigDecimal.class, params);
        log.info("Formato136QueryService resultado metric={} valor={}", metric, value);
        return value == null ? BigDecimal.ZERO : value;
    }

    private String sqlFondoEstandar() {
        return """
                SELECT a.Codigo_Entidad AS codigo_entidad,
                       'ENT_' || TRIM(CAST(a.Codigo_Entidad AS VARCHAR(20))) AS clave_reporte,
                       COALESCE(SUM(e.valor), 0) / 1000000 AS valor_mm_cop
                FROM PROD_DWH_CONSULTA.ENTIDADES a
                INNER JOIN PROD_DWH_CONSULTA.NEGFID_INSUMO_ENTIDAD e ON e.ent_id = a.ent_id
                INNER JOIN PROD_DWH_CONSULTA.TIEMPO b ON e.tie_id = b.tie_id
                INNER JOIN PROD_DWH_CONSULTA.PATRIMONIOS_AUTONOMOS c ON e.paau_id = c.paau_id
                INNER JOIN PROD_DWH_CONSULTA.NEGFID_INSUMOS d ON d.inf_id = e.inf_id
                WHERE c.tipo_patrimonio = 6
                  AND c.codigo_patrimonio = ?
                  AND d.nivel1 = 136
                  AND d.nivel2 = 2
                  AND d.nivel3 = 4
                  AND d.nivel4 = 305
                  AND a.tipo_entidad = 23
                  AND e.valor <> 0
                  AND b.fecha = ?
                GROUP BY 1, 2
                ORDER BY 1
                """;
    }
    private String sqlFondoModerado() {
        return """
                SELECT a.Codigo_Entidad AS codigo_entidad,
                       CASE
                           WHEN a.Codigo_Entidad = 9 AND c.Codigo_Patrimonio = 8000 THEN 'SKANDIA_ALT'
                           ELSE 'ENT_' || TRIM(CAST(a.Codigo_Entidad AS VARCHAR(20)))
                       END AS clave_reporte,
                       COALESCE(SUM(e.valor), 0) / 1000000 AS valor_mm_cop
                FROM PROD_DWH_CONSULTA.ENTIDADES a
                INNER JOIN PROD_DWH_CONSULTA.NEGFID_INSUMO_ENTIDAD e ON e.ent_id = a.ent_id
                INNER JOIN PROD_DWH_CONSULTA.TIEMPO b ON e.tie_id = b.tie_id
                INNER JOIN PROD_DWH_CONSULTA.PATRIMONIOS_AUTONOMOS c ON e.paau_id = c.paau_id
                INNER JOIN PROD_DWH_CONSULTA.NEGFID_INSUMOS d ON d.inf_id = e.inf_id
                WHERE c.tipo_patrimonio = 6
                  AND c.Codigo_Patrimonio IN (1000, 8000)
                  AND d.nivel1 = 136
                  AND d.nivel2 = 2
                  AND d.nivel3 = 4
                  AND d.nivel4 = 305
                  AND a.tipo_entidad = 23
                  AND e.valor <> 0
                  AND b.fecha = ?
                GROUP BY 1, 2
                ORDER BY 1
                """;
    }
    private String sqlAportesRecibidos() {
        return """
                SELECT COALESCE(SUM(e.valor) / 1000000, 0) AS Valor_Total
                FROM prod_dwh_consulta.entidades a,
                     prod_dwh_consulta.tiempo b,
                     prod_dwh_consulta.patrimonios_autonomos c,
                     prod_dwh_consulta.negfid_insumos d,
                     prod_dwh_consulta.negfid_insumo_entidad e
                WHERE d.inf_id = e.inf_id
                  AND e.ent_id = a.ent_id
                  AND e.tie_id = b.tie_id
                  AND e.paau_id = c.paau_id
                  AND c.tipo_patrimonio = 6
                  AND c.codigo_patrimonio = 1000
                  AND d.nivel1 = 136
                  AND d.nivel2 = 2
                  AND d.nivel3 = 4
                  AND d.nivel4 = 10
                  AND a.tipo_entidad = 23
                  AND e.valor <> 0
                  AND b.fecha BETWEEN ? AND ?
                """;
    }
}
