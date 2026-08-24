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

        log.info("Fondo administrado desde Teradata NEGFID_INSUMO_ENTIDAD fechaCorte={} totalMmCop={} proteccionPorvenirMmCop={} concentracionPct={}",
                fechaCorte, totalMmCop, proteccionPorvenirMmCop, concentracionProteccionPorvenirPct);
        return new FondoAdministrado(totalMmCop, concentracionProteccionPorvenirPct);
    }

    private String sqlFondoPorEntidad() {
        return """
                SELECT a.Codigo_Entidad AS CODIGO_ENTIDAD,
                       COALESCE(SUM(e.valor), 0) / 1000000 AS VALOR_MM_COP
                FROM PROD_DWH_CONSULTA.ENTIDADES a
                INNER JOIN PROD_DWH_CONSULTA.NEGFID_INSUMO_ENTIDAD e ON e.ent_id = a.ent_id
                INNER JOIN PROD_DWH_CONSULTA.TIEMPO b ON e.tie_id = b.tie_id
                INNER JOIN PROD_DWH_CONSULTA.PATRIMONIOS_AUTONOMOS c ON e.paau_id = c.paau_id
                INNER JOIN PROD_DWH_CONSULTA.NEGFID_INSUMOS d ON d.inf_id = e.inf_id
                WHERE c.tipo_patrimonio = 6
                  AND c.codigo_patrimonio IN (1000, 5000, 6000, 7000, 8000)
                  AND d.nivel1 = 136
                  AND d.nivel2 = 2
                  AND d.nivel3 = 4
                  AND d.nivel4 = 305
                  AND a.tipo_entidad = 23
                  AND e.valor <> 0
                  AND b.fecha = ?
                GROUP BY a.Codigo_Entidad
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
