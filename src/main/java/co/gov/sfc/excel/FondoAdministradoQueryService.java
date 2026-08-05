package co.gov.sfc.excel;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.jdbc.core.JdbcTemplate;
import org.springframework.stereotype.Component;

import java.math.BigDecimal;
import java.math.RoundingMode;
import java.sql.Date;
import java.time.LocalDate;
import java.util.List;

@Component
public class FondoAdministradoQueryService {

    private static final Logger log = LoggerFactory.getLogger(FondoAdministradoQueryService.class);
    private final JdbcTemplate jdbcTemplate;

    public FondoAdministradoQueryService(JdbcTemplate jdbcTemplate) {
        this.jdbcTemplate = jdbcTemplate;
    }

    public FondoAdministrado leer(LocalDate fechaCorte) {
        Date fecha = Date.valueOf(fechaCorte);
        List<FondoPorEntidad> valores = jdbcTemplate.query(sqlFondoPorEntidad(), (rs, rowNum) ->
                new FondoPorEntidad(
                        rs.getInt("CODIGO_ENTIDAD"),
                        nvl(rs.getBigDecimal("VALOR_MM_COP"))),
                fecha);

        BigDecimal totalMmCop = valores.stream()
                .map(FondoPorEntidad::valorMmCop)
                .reduce(BigDecimal.ZERO, BigDecimal::add);
        BigDecimal proteccionPorvenirMmCop = valores.stream()
                .filter(v -> v.codigoEntidad() == 2 || v.codigoEntidad() == 3)
                .map(FondoPorEntidad::valorMmCop)
                .reduce(BigDecimal.ZERO, BigDecimal::add);
        BigDecimal concentracionProteccionPorvenirPct = totalMmCop.signum() == 0
                ? BigDecimal.ZERO
                : proteccionPorvenirMmCop
                        .divide(totalMmCop, 10, RoundingMode.HALF_UP)
                        .multiply(BigDecimal.valueOf(100));

        log.info("Fondo administrado desde Teradata fechaCorte={} totalMmCop={} proteccionPorvenirMmCop={} concentracionPct={}",
                fechaCorte, totalMmCop, proteccionPorvenirMmCop, concentracionProteccionPorvenirPct);
        return new FondoAdministrado(totalMmCop, concentracionProteccionPorvenirPct);
    }

    private String sqlFondoPorEntidad() {
        return """
                SELECT e.Codigo_Entidad AS CODIGO_ENTIDAD,
                       COALESCE(SUM(COALESCE(eip.Saldo_Sincierre_Total_Moneda_0, 0)), 0) / 1000000 AS VALOR_MM_COP
                FROM PROD_DWH_CONSULTA.ESTFIN_INDIV_PA eip
                INNER JOIN PROD_DWH_CONSULTA.ENTIDADES e
                        ON eip.Ent_ID = e.Ent_ID
                INNER JOIN PROD_DWH_CONSULTA.PATRIMONIOS_AUTONOMOS pa
                        ON eip.Paau_ID = pa.Paau_ID
                INNER JOIN PROD_DWH_CONSULTA.TIEMPO t
                        ON eip.Tie_ID = t.Tie_ID
                INNER JOIN PROD_DWH_CONSULTA.PUC p
                        ON eip.Puc_ID = p.Puc_ID
                WHERE eip.Tipo_Informe = 17
                  AND e.Tipo_Entidad = 23
                  AND e.Estado = 1
                  AND pa.Tipo_Patrimonio = 6
                  AND p.Codigo = 100000
                  AND t.Fecha = ?
                  AND (
                        pa.Codigo_Patrimonio IN (1000, 5000, 6000, 7000, 8000)
                        OR (e.Codigo_Entidad = 9 AND pa.Codigo_Patrimonio = 4)
                  )
                GROUP BY e.Codigo_Entidad
                """;
    }

    private BigDecimal nvl(BigDecimal value) {
        return value == null ? BigDecimal.ZERO : value;
    }

    private record FondoPorEntidad(int codigoEntidad, BigDecimal valorMmCop) {}

    public record FondoAdministrado(
            BigDecimal totalMmCop,
            BigDecimal concentracionProteccionPorvenirPct
    ) {}
}
