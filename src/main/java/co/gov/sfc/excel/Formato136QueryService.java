package co.gov.sfc.excel;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.jdbc.core.JdbcTemplate;
import org.springframework.stereotype.Component;

import java.math.BigDecimal;
import java.sql.Date;
import java.time.LocalDate;
import java.util.HashMap;
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
        log.info("Formato136QueryService ejecutando metric=colombia_por_fondo_entidad fechaCorte={} sql=\"{}\"",
                fechaCorte,
                sqlColombiaPorFondoEntidad().replace("\n", " ").replaceAll("\\s+", " ").trim());
        jdbcTemplate.query(sqlColombiaPorFondoEntidad(), rs -> {
            int codigoEntidad = rs.getInt("codigo_entidad");
            int codigoPatrimonio = rs.getInt("codigo_patrimonio");
            BigDecimal valor = rs.getBigDecimal("valor_total");
            putColombiaValue(out, codigoPatrimonio, codigoEntidad, valor == null ? BigDecimal.ZERO : valor);
        }, Date.valueOf(fechaCorte));
        log.info("Formato136QueryService resultado metric=colombia_por_fondo_entidad fechaCorte={} valores={}", fechaCorte, out);
        return out;
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

    private void putColombiaValue(Map<String, BigDecimal> out, int codigoPatrimonio, int codigoEntidad, BigDecimal valor) {
        String entidad = switch (codigoEntidad) {
            case 10 -> "colf";
            case 3 -> "porv";
            case 2 -> "prot";
            case 9 -> "sk";
            default -> null;
        };
        if (entidad == null) {
            return;
        }
        String fondo = switch (codigoPatrimonio) {
            case 1000 -> "mod";
            case 5000 -> "con";
            case 6000 -> "mr";
            case 7000 -> "rp";
            case 8000 -> "alt";
            default -> null;
        };
        if (fondo == null) {
            return;
        }
        if ("alt".equals(fondo)) {
            out.merge("mod_alt", valor, BigDecimal::add);
        } else {
            out.merge(fondo + "_" + entidad, valor, BigDecimal::add);
        }
    }

    private String sqlColombiaPorFondoEntidad() {
        return """
                SELECT a.codigo_entidad,
                       c.codigo_patrimonio,
                       COALESCE(SUM(e.valor) / 1000000, 0) AS valor_total
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
                  AND c.codigo_patrimonio IN (1000, 5000, 6000, 7000, 8000)
                  AND d.nivel1 = 136
                  AND d.nivel2 = 2
                  AND d.nivel3 = 4
                  AND d.nivel4 = 10
                  AND a.tipo_entidad = 23
                  AND a.codigo_entidad IN (2, 3, 9, 10)
                  AND e.valor <> 0
                  AND b.fecha = ?
                GROUP BY a.codigo_entidad, c.codigo_patrimonio
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
