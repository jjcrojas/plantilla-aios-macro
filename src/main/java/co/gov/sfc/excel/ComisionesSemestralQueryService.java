package co.gov.sfc.excel;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.jdbc.core.JdbcTemplate;
import org.springframework.stereotype.Component;

import java.math.BigDecimal;
import java.sql.Date;
import java.time.LocalDate;

@Component
public class ComisionesSemestralQueryService {

    private static final Logger log = LoggerFactory.getLogger(ComisionesSemestralQueryService.class);
    private final JdbcTemplate jdbcTemplate;

    public ComisionesSemestralQueryService(JdbcTemplate jdbcTemplate) {
        this.jdbcTemplate = jdbcTemplate;
    }

    public BigDecimal leer411500(LocalDate fechaCorte, BigDecimal trm) {
        return leerCuenta(fechaCorte, trm, 411500);
    }

    public BigDecimal leerCuenta(LocalDate fechaCorte, BigDecimal trm, int codigoCuenta) {
        if (trm == null || trm.signum() == 0) {
            throw new IllegalArgumentException("La TRM debe ser distinta de cero para consultar la cuenta " + codigoCuenta);
        }
        LocalDate cierreAnterior = LocalDate.of(fechaCorte.getYear() - 1, 12, 31);
        LocalDate corteAnterior = fechaCorte.minusYears(1);
        Object[] params = {
                Date.valueOf(fechaCorte), Date.valueOf(cierreAnterior), Date.valueOf(corteAnterior),
                trm,
                Date.valueOf(fechaCorte), Date.valueOf(cierreAnterior), Date.valueOf(corteAnterior),
                codigoCuenta
        };
        log.info("Semestral query cuenta={} fechaCorte={} cierreAnterior={} corteAnterior={} trm={} sql=\"{}\"",
                codigoCuenta, fechaCorte, cierreAnterior, corteAnterior, trm,
                sqlCuenta().replace("\n", " ").replaceAll("\\s+", " ").trim());
        BigDecimal value = jdbcTemplate.queryForObject(sqlCuenta(), BigDecimal.class, params);
        BigDecimal result = value == null ? BigDecimal.ZERO : value;
        log.info("Semestral resultado cuenta={} valorUsd={} fechaCorte={} trm={}", codigoCuenta, result, fechaCorte, trm);
        return result;
    }

    public BigDecimal leerGastosOperativos(LocalDate fechaCorte, BigDecimal trm) {
        if (trm == null || trm.signum() == 0) {
            throw new IllegalArgumentException("La TRM debe ser distinta de cero para calcular la fila 52");
        }
        LocalDate cierreAnterior = LocalDate.of(fechaCorte.getYear() - 1, 12, 31);
        LocalDate corteAnterior = fechaCorte.minusYears(1);
        Object[] params = {
                Date.valueOf(fechaCorte), Date.valueOf(cierreAnterior), Date.valueOf(corteAnterior),
                Date.valueOf(fechaCorte), Date.valueOf(cierreAnterior), Date.valueOf(corteAnterior),
                trm,
                Date.valueOf(fechaCorte), Date.valueOf(cierreAnterior), Date.valueOf(corteAnterior)
        };
        log.info("Semestral fila 52 query gastosOperativos fechaCorte={} cierreAnterior={} corteAnterior={} trm={} sql=\"{}\"",
                fechaCorte, cierreAnterior, corteAnterior, trm,
                sqlGastosOperativos().replace("\n", " ").replaceAll("\\s+", " ").trim());
        BigDecimal value = jdbcTemplate.queryForObject(sqlGastosOperativos(), BigDecimal.class, params);
        BigDecimal result = value == null ? BigDecimal.ZERO : value;
        log.info("Semestral fila 52 resultado gastosOperativosUsd={} fechaCorte={} trm={}", result, fechaCorte, trm);
        return result;
    }

    String sqlGastosOperativos() {
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
                  AND ei.Tipo_Informe = 0
                  AND t.Fecha IN (?, ?, ?)
                  AND p.Codigo IN (510000, 510300, 510400, 510600, 510700, 510800,
                                   512500, 512800, 512900, 513900)
                """;
    }
    String sqlCuenta() {
        return """
                SELECT COALESCE(SUM(
                           (CASE WHEN t.Fecha = ? THEN ei.Saldo_Sincierre_Total_Moneda_0 ELSE 0 END)
                         + (CASE WHEN t.Fecha = ? THEN ei.Saldo_Sincierre_Total_Moneda_0 ELSE 0 END)
                         - (CASE WHEN t.Fecha = ? THEN ei.Saldo_Sincierre_Total_Moneda_0 ELSE 0 END)
                       ), 0) / 1000000 / ? AS RESULTADO
                FROM PROD_DWH_CONSULTA.ESTFIN_INDIV ei
                INNER JOIN PROD_DWH_CONSULTA.ENTIDADES e ON ei.Ent_ID = e.Ent_ID
                INNER JOIN PROD_DWH_CONSULTA.TIEMPO t ON ei.Tie_ID = t.Tie_ID
                INNER JOIN PROD_DWH_CONSULTA.PUC p ON ei.Puc_ID = p.Puc_ID
                WHERE e.Tipo_Entidad = 23
                  AND ei.Tipo_Informe = 0
                  AND t.Fecha IN (?, ?, ?)
                  AND p.Codigo = ?
                """;
    }
}
