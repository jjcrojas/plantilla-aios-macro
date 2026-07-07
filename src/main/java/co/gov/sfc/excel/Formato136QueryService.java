package co.gov.sfc.excel;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.jdbc.core.JdbcTemplate;
import org.springframework.stereotype.Component;

import java.math.BigDecimal;
import java.sql.Date;
import java.time.LocalDate;

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

    private BigDecimal scalar(String metric, String sql, Object... params) {
        log.info("Formato136QueryService ejecutando metric={} params={} sql=\"{}\"",
                metric,
                java.util.Arrays.toString(params),
                sql.replace("\n", " ").replaceAll("\\s+", " ").trim());
        BigDecimal value = jdbcTemplate.queryForObject(sql, BigDecimal.class, params);
        log.info("Formato136QueryService resultado metric={} valor={}", metric, value);
        return value == null ? BigDecimal.ZERO : value;
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
