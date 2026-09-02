package co.gov.sfc.excel;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.jdbc.core.JdbcTemplate;
import org.springframework.stereotype.Component;

import java.math.BigDecimal;
import java.sql.Date;
import java.time.LocalDate;
import java.util.LinkedHashSet;
import java.util.List;

@Component
public class BalanceContableQueryService {

    static final int CUENTA_ACTIVO = 100000;
    static final int CUENTA_PASIVO = 200000;
    static final int CUENTA_PATRIMONIO = 300000;

    private static final Logger log = LoggerFactory.getLogger(BalanceContableQueryService.class);
    private final JdbcTemplate jdbcTemplate;

    public BalanceContableQueryService(JdbcTemplate jdbcTemplate) {
        this.jdbcTemplate = jdbcTemplate;
    }

    public BalanceContable leer(LocalDate fechaCorte) {
        List<String> administradoras = leerAdministradorasVigentes();
        BalanceContable balance = new BalanceContable(
                leerCuenta(fechaCorte, CUENTA_ACTIVO),
                leerCuenta(fechaCorte, CUENTA_PASIVO),
                leerCuenta(fechaCorte, CUENTA_PATRIMONIO),
                administradoras
        );
        log.info("Balance contable consultado fechaCorte={} activoMmCop={} pasivoMmCop={} patrimonioMmCop={} administradorasVigentes={}",
                fechaCorte, balance.activoMmCop(), balance.pasivoMmCop(), balance.patrimonioMmCop(),
                balance.administradorasVigentes());
        return balance;
    }

    BigDecimal leerCuenta(LocalDate fechaCorte, int codigoCuenta) {
        BigDecimal value = jdbcTemplate.queryForObject(
                sqlCuenta(), BigDecimal.class, codigoCuenta, Date.valueOf(fechaCorte));
        BigDecimal result = value == null ? BigDecimal.ZERO : value;
        log.info("Saldo contable consultado fechaCorte={} cuenta={} valorMmCop={}",
                fechaCorte, codigoCuenta, result);
        return result;
    }

    List<String> leerAdministradorasVigentes() {
        List<String> nombres = jdbcTemplate.queryForList(sqlAdministradorasVigentes(), String.class);
        LinkedHashSet<String> sinComillas = new LinkedHashSet<>();
        for (String nombre : nombres) {
            String limpio = quitarComillas(nombre);
            if (!limpio.isBlank()) {
                sinComillas.add(limpio);
            }
        }
        return List.copyOf(sinComillas);
    }

    String sqlCuenta() {
        return """
                SELECT COALESCE(SUM(ef.Saldo_Sincierre_Total_Moneda_0), 0) / 1000000 AS VALOR_MM_COP
                FROM PROD_DWH_CONSULTA.ESTFIN_INDIV ef
                INNER JOIN PROD_DWH_CONSULTA.TIEMPO t ON t.Tie_ID = ef.Tie_ID
                INNER JOIN PROD_DWH_CONSULTA.ENTIDADES e ON e.Ent_ID = ef.Ent_ID
                INNER JOIN PROD_DWH_CONSULTA.PUC p ON p.Puc_ID = ef.Puc_ID
                WHERE e.Tipo_Entidad = 23
                  AND ef.Tipo_Informe = 0
                  AND p.Codigo = ?
                  AND t.Fecha = ?
                """;
    }

    String sqlAdministradorasVigentes() {
        return """
                SELECT DISTINCT TRIM(e.Nombre_Entidad)
                FROM PROD_DWH_CONSULTA.ENTIDADES e
                WHERE e.Tipo_Entidad = 23
                  AND e.Estado = 1
                ORDER BY 1
                """;
    }

    private String quitarComillas(String value) {
        if (value == null) return "";
        return value
                .replace("\"", "")
                .replace("'", "")
                .replace("“", "")
                .replace("”", "")
                .replace("‘", "")
                .replace("’", "")
                .trim();
    }

    public record BalanceContable(
            BigDecimal activoMmCop,
            BigDecimal pasivoMmCop,
            BigDecimal patrimonioMmCop,
            List<String> administradorasVigentes
    ) {
        public BigDecimal numeroAdministradorasVigentes() {
            return BigDecimal.valueOf(administradorasVigentes.size());
        }
    }
}
