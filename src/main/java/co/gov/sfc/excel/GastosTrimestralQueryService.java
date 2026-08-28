package co.gov.sfc.excel;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.jdbc.core.JdbcTemplate;
import org.springframework.stereotype.Component;

import java.math.BigDecimal;
import java.sql.Date;
import java.time.LocalDate;
import java.util.LinkedHashMap;
import java.util.Map;

@Component
public class GastosTrimestralQueryService {

    private static final Logger log = LoggerFactory.getLogger(GastosTrimestralQueryService.class);
    private final JdbcTemplate jdbcTemplate;

    public GastosTrimestralQueryService(JdbcTemplate jdbcTemplate) {
        this.jdbcTemplate = jdbcTemplate;
    }

    public Map<String, BigDecimal> leerGastosUsd(LocalDate fechaCorte, BigDecimal trm) {
        Map<String, BigDecimal> gastos = new LinkedHashMap<>();
        gastos.put("colf", leerGastoPorEntidad(fechaCorte, trm, 10));
        gastos.put("porv", leerGastoPorEntidad(fechaCorte, trm, 3));
        gastos.put("prot", leerGastoPorEntidad(fechaCorte, trm, 2));
        gastos.put("sk", leerGastoPorEntidad(fechaCorte, trm, 9));
        log.info("Gastos trimestrales consultados en Teradata fechaCorte={} trm={} valoresUsd={}",
                fechaCorte, trm, gastos);
        return gastos;
    }

    BigDecimal leerGastoPorEntidad(LocalDate fechaCorte, BigDecimal trm, int codigoEntidad) {
        if (trm == null || trm.signum() == 0) {
            throw new IllegalArgumentException("La TRM debe ser distinta de cero para consultar gastos trimestrales");
        }
        LocalDate cierreAnterior = LocalDate.of(fechaCorte.getYear() - 1, 12, 31);
        LocalDate corteAnterior = fechaCorte.minusYears(1);
        Object[] params = {
                Date.valueOf(fechaCorte), Date.valueOf(cierreAnterior), Date.valueOf(corteAnterior),
                Date.valueOf(fechaCorte), Date.valueOf(cierreAnterior), Date.valueOf(corteAnterior),
                trm,
                codigoEntidad,
                Date.valueOf(fechaCorte), Date.valueOf(cierreAnterior), Date.valueOf(corteAnterior)
        };
        log.info("Gastos trimestrales query fechaCorte={} cierreAnterior={} corteAnterior={} trm={} codigoEntidad={} sql=\"{}\"",
                fechaCorte, cierreAnterior, corteAnterior, trm, codigoEntidad,
                sqlGastosPorEntidad().replace("\n", " ").replaceAll("\\s+", " ").trim());
        BigDecimal value = jdbcTemplate.queryForObject(sqlGastosPorEntidad(), BigDecimal.class, params);
        BigDecimal result = value == null ? BigDecimal.ZERO : value;
        log.info("Gastos trimestrales resultado codigoEntidad={} valorUsd={} fechaCorte={} trm={}",
                codigoEntidad, result, fechaCorte, trm);
        return result;
    }

    String sqlGastosPorEntidad() {
        return """
                SELECT COALESCE(SUM(
                    CASE
                       WHEN p.Codigo = 510000 THEN
                              (CASE WHEN t.Fecha = ? THEN ei.Saldo_Sincierre_Total_Moneda_0 ELSE 0 END)
                            + (CASE WHEN t.Fecha = ? THEN ei.Saldo_Sincierre_Total_Moneda_0 ELSE 0 END)
                            - (CASE WHEN t.Fecha = ? THEN ei.Saldo_Sincierre_Total_Moneda_0 ELSE 0 END)
                       WHEN p.Codigo IN (510300, 510400, 510600, 510700, 510800,
                                         512500, 512800, 512900, 513900) THEN -(
                              (CASE WHEN t.Fecha = ? THEN ei.Saldo_Sincierre_Total_Moneda_0 ELSE 0 END)
                            + (CASE WHEN t.Fecha = ? THEN ei.Saldo_Sincierre_Total_Moneda_0 ELSE 0 END)
                            - (CASE WHEN t.Fecha = ? THEN ei.Saldo_Sincierre_Total_Moneda_0 ELSE 0 END)
                       )
                       ELSE 0
                    END
                ), 0) / 1000000 / ? AS RESULTADO
                FROM PROD_DWH_CONSULTA.ESTFIN_INDIV ei
                INNER JOIN PROD_DWH_CONSULTA.ENTIDADES e ON ei.Ent_ID = e.Ent_ID
                INNER JOIN PROD_DWH_CONSULTA.TIEMPO t ON ei.Tie_ID = t.Tie_ID
                INNER JOIN PROD_DWH_CONSULTA.PUC p ON ei.Puc_ID = p.Puc_ID
                WHERE e.Tipo_Entidad = 23
                  AND e.Codigo_Entidad = ?
                  AND ei.Tipo_Informe = 0
                  AND t.Fecha IN (?, ?, ?)
                  AND p.Codigo IN (510000, 510300, 510400, 510600, 510700, 510800,
                                   512500, 512800, 512900, 513900)
                """;
    }
}
