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
    private static final BigDecimal MIL = BigDecimal.valueOf(1_000);

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
                out.merge(fondo + "_" + entidad, valorMillonesCop(rs.getBigDecimal("valor_miles")), BigDecimal::add);
            }
        }, codigoPatrimonio, Date.valueOf(fechaCorte));
    }

    private void leerFondoModerado(Map<String, BigDecimal> out, LocalDate fechaCorte) {
        String sql = sqlFondoModerado();
        logQueryColombia("mod", fechaCorte, sql, 1000, 4, 8000);
        jdbcTemplate.query(sql, rs -> {
            String claveReporte = rs.getString("clave_reporte");
            BigDecimal valor = valorMillonesCop(rs.getBigDecimal("valor_miles"));
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

    private BigDecimal valorMillonesCop(BigDecimal valorMiles) {
        // ESTFIN_INDIV_PA devuelve miles de COP; TrimestralDataReader divide después por la TRM.
        // Esta segunda división por 1.000 deja el saldo en millones de COP, igual que la macro VBA.
        return valorMiles == null ? BigDecimal.ZERO : valorMiles.divide(MIL);
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
                SELECT e.Codigo_Entidad AS codigo_entidad,
                       'ENT_' || TRIM(CAST(e.Codigo_Entidad AS VARCHAR(20))) AS clave_reporte,
                       SUM(eip.Saldo_Sincierre_Total_Moneda_0) / 1000 AS valor_miles
                FROM PROD_DWH_CONSULTA.ESTFIN_INDIV_PA eip
                INNER JOIN PROD_DWH_CONSULTA.ENTIDADES e ON eip.Ent_ID = e.Ent_ID
                INNER JOIN PROD_DWH_CONSULTA.PATRIMONIOS_AUTONOMOS pa ON eip.Paau_ID = pa.Paau_ID
                INNER JOIN PROD_DWH_CONSULTA.TIEMPO t ON eip.Tie_ID = t.Tie_ID
                INNER JOIN PROD_DWH_CONSULTA.PUC p ON eip.Puc_ID = p.Puc_ID
                WHERE eip.Tipo_Informe = 17
                  AND e.Tipo_Entidad = 23
                  AND e.Estado = 1
                  AND pa.Tipo_Patrimonio = 6
                  AND pa.Codigo_Patrimonio = ?
                  AND p.Codigo = 100000
                  AND t.Fecha = ?
                GROUP BY 1, 2
                ORDER BY 1
                """;
    }

    private String sqlFondoModerado() {
        return """
                SELECT e.Codigo_Entidad AS codigo_entidad,
                       CASE
                           WHEN e.Codigo_Entidad = 9
                            AND pa.Tipo_Patrimonio = 6
                            AND pa.Codigo_Patrimonio IN (4, 8000) THEN 'SKANDIA_ALT'
                           ELSE 'ENT_' || TRIM(CAST(e.Codigo_Entidad AS VARCHAR(20)))
                       END AS clave_reporte,
                       SUM(eip.Saldo_Sincierre_Total_Moneda_0) / 1000 AS valor_miles
                FROM PROD_DWH_CONSULTA.ESTFIN_INDIV_PA eip
                INNER JOIN PROD_DWH_CONSULTA.ENTIDADES e ON eip.Ent_ID = e.Ent_ID
                INNER JOIN PROD_DWH_CONSULTA.PATRIMONIOS_AUTONOMOS pa ON eip.Paau_ID = pa.Paau_ID
                INNER JOIN PROD_DWH_CONSULTA.TIEMPO t ON eip.Tie_ID = t.Tie_ID
                INNER JOIN PROD_DWH_CONSULTA.PUC p ON eip.Puc_ID = p.Puc_ID
                WHERE eip.Tipo_Informe = 17
                  AND e.Tipo_Entidad = 23
                  AND e.Estado = 1
                  AND pa.Tipo_Patrimonio = 6
                  AND (pa.Codigo_Patrimonio = 1000
                       OR (e.Codigo_Entidad = 9 AND pa.Codigo_Patrimonio IN (4, 8000)))
                  AND p.Codigo = 100000
                  AND t.Fecha = ?
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
