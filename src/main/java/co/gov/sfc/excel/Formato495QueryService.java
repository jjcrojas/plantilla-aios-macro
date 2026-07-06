package co.gov.sfc.excel;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.jdbc.core.JdbcTemplate;
import org.springframework.stereotype.Component;

import java.math.BigDecimal;
import java.sql.Date;
import java.time.LocalDate;

@Component
public class Formato495QueryService {

    private static final Logger log = LoggerFactory.getLogger(Formato495QueryService.class);

    private final JdbcTemplate jdbcTemplate;

    public Formato495QueryService(JdbcTemplate jdbcTemplate) {
        this.jdbcTemplate = jdbcTemplate;
    }

    public PensionadosResumen leerResumen(LocalDate fechaCorte) {
        PensionadosResumen resumen = new PensionadosResumen(
                leerTotal(fechaCorte),
                leerInvalidez(fechaCorte),
                leerVejez(fechaCorte),
                leerSobrevivencia(fechaCorte)
        );
        log.info("Formato495QueryService resultado fechaCorte={} resumen={}", fechaCorte, resumen);
        return resumen;
    }

    public BigDecimal leerTotal(LocalDate fechaCorte) {
        return scalar("total_pensionados", sqlTotal(), Date.valueOf(fechaCorte));
    }

    public BigDecimal leerInvalidez(LocalDate fechaCorte) {
        return scalar("pensionados_invalidez", sqlInvalidez(), Date.valueOf(fechaCorte));
    }

    public BigDecimal leerVejez(LocalDate fechaCorte) {
        return scalar("pensionados_vejez", sqlVejez(), Date.valueOf(fechaCorte));
    }

    public BigDecimal leerSobrevivencia(LocalDate fechaCorte) {
        return scalar("pensionados_sobrevivencia", sqlSobrevivencia(), Date.valueOf(fechaCorte));
    }

    private BigDecimal scalar(String metric, String sql, Object... params) {
        log.info("Formato495QueryService ejecutando metric={} params={} sql=\"{}\"",
                metric,
                java.util.Arrays.toString(params),
                sql.replace("\n", " ").replaceAll("\\s+", " ").trim());
        BigDecimal value = jdbcTemplate.queryForObject(sql, BigDecimal.class, params);
        log.info("Formato495QueryService resultado metric={} valor={}", metric, value);
        return value == null ? BigDecimal.ZERO : value;
    }

    private String filtroBase() {
        return """
                FROM PROD_DWH_CONSULTA.S9_FORMATO_495
                WHERE FECHA_CORTE = ?
                  AND UNIDAD_CAPTURA = 1
                  AND RENGLON = 200
                """;
    }

    private String sqlTotal() {
        return """
                SELECT COALESCE(
                    SUM(COALESCE(RTR_PRGRMD_V_H,0))
                  + SUM(COALESCE(RTR_PRGRMD_V_M,0))
                  + SUM(COALESCE(RTR_PRGRMD_I_H,0))
                  + SUM(COALESCE(RTR_PRGRMD_I_M,0))
                  + SUM(COALESCE(RTR_PRGRMD_S_H,0))
                  + SUM(COALESCE(RTR_PRGRMD_S_M,0))
                  + SUM(COALESCE(RTR_PRGRMD_RNT_VTLC_DIF_V_H,0))
                  + SUM(COALESCE(RTR_PRGRMD_RNT_VTLC_DIF_I_H,0))
                  + SUM(COALESCE(RTR_PRGRMD_RNT_VTLC_DIF_I_M,0))
                  + SUM(COALESCE(RTR_PRGRMD_RNT_VTLC_DIF_I,0))
                  + SUM(COALESCE(RTR_PRGRMD_RNT_VTLC_DIF_S_H,0))
                  + SUM(COALESCE(RTR_PRGRMD_RNT_VTLC_DIF_S_M,0))
                  + SUM(COALESCE(PNSNS_FLL_JDCL_V_H,0))
                  + SUM(COALESCE(PNSNS_FLL_JDCL_V_M,0))
                  + SUM(COALESCE(PNSNS_FLL_JDCL_I_H,0))
                  + SUM(COALESCE(PNSNS_FLL_JDCL_I_M,0))
                  + SUM(COALESCE(PNSNS_FLL_JDCL_S_H,0))
                  + SUM(COALESCE(PNSNS_FLL_JDCL_S_M,0))
                  + SUM(COALESCE(RNT_VTLC_INMDT_V_H,0))
                  + SUM(COALESCE(RNT_VTLC_INMDT_V_M,0))
                  + SUM(COALESCE(RNT_VTLC_INMDT_I_H,0))
                  + SUM(COALESCE(RNT_VTLC_INMDT_I_M,0))
                  + SUM(COALESCE(RNT_VTLC_INMDT_S_H,0))
                  + SUM(COALESCE(RNT_VTLC_INMDT_S_M,0))
                  + SUM(COALESCE(RTR_PRGRMD_SIN_NGCCN_BONO_V_H,0))
                  + SUM(COALESCE(RTR_PRGRMD_SIN_NGCCN_BONO_V_M,0))
                  + SUM(COALESCE(RNT_TMP_CRT_VTLC_DFM_CRT_V_H,0))
                  + SUM(COALESCE(RNT_TMP_CRT_VTLC_DFM_CRT_V_M,0))
                  + SUM(COALESCE(RNT_TMP_CRT_VTLC_DFM_CRT_I_H,0))
                  + SUM(COALESCE(RNT_TMP_CRT_VTLC_DFM_CRT_I_M,0))
                  + SUM(COALESCE(RNT_TMP_CRT_VTLC_DFM_CRT_S_H,0))
                  + SUM(COALESCE(RNT_TMP_CRT_VTLC_DFM_CRT_S_M,0))
                  + SUM(COALESCE(RNT_TMP_VRBL_RNT_VTLC_DIF_V_H,0))
                  + SUM(COALESCE(RNT_TMP_VRBL_RNT_VTLC_DIF_V_M,0))
                  + SUM(COALESCE(RNT_TMP_VRBL_RNT_VTLC_DIF_I_H,0))
                  + SUM(COALESCE(RNT_TMP_VRBL_RNT_VTLC_DIF_I_M,0))
                  + SUM(COALESCE(RNT_TMP_VRBL_RNT_VTLC_DIF_I,0))
                  + SUM(COALESCE(RNT_TMP_VRBL_RNT_VTLC_DIF_S_H,0))
                  + SUM(COALESCE(RNT_TMP_VRBL_RNT_VTLC_INMD_V_H,0))
                  + SUM(COALESCE(RNT_TMP_VRBL_RNT_VTLC_INMD_V_M,0))
                  + SUM(COALESCE(RNT_TMP_VRBL_RNT_VTLC_INMD_I_H,0))
                  + SUM(COALESCE(RNT_TMP_VRBL_RNT_VTLC_INMD_I_M,0))
                  + SUM(COALESCE(RNT_TMP_VRBL_RNT_VTLC_INMDT_I,0))
                  + SUM(COALESCE(RNT_TMP_VRBL_RNT_VTLC_INMD_S_H,0))
                  + SUM(COALESCE(RNT_TMP_VRBL_RNT_VTLC_INMD_S_M,0)), 0) AS TOTAL
                """ + filtroBase();
    }

    private String sqlInvalidez() {
        return """
                SELECT COALESCE(
                      SUM(COALESCE(RNT_TMP_VRBL_RNT_VTLC_INMD_I_M,0))
                    + SUM(COALESCE(RNT_TMP_VRBL_RNT_VTLC_INMD_I_H,0))
                    + SUM(COALESCE(RNT_TMP_VRBL_RNT_VTLC_DIF_I_M,0))
                    + SUM(COALESCE(RNT_TMP_VRBL_RNT_VTLC_DIF_I_H,0))
                    + SUM(COALESCE(RNT_VTLC_INMDT_I_M,0))
                    + SUM(COALESCE(RNT_VTLC_INMDT_I_H,0))
                    + SUM(COALESCE(RNT_TMP_CRT_VTLC_DFM_CRT_I_M,0))
                    + SUM(COALESCE(RNT_TMP_CRT_VTLC_DFM_CRT_I_H,0))
                    + SUM(COALESCE(PNSNS_FLL_JDCL_I_M,0))
                    + SUM(COALESCE(PNSNS_FLL_JDCL_I_H,0))
                    + SUM(COALESCE(RTR_PRGRMD_RNT_VTLC_DIF_I_M,0))
                    + SUM(COALESCE(RTR_PRGRMD_RNT_VTLC_DIF_I_H,0))
                    + SUM(COALESCE(RTR_PRGRMD_I_M,0))
                    + SUM(COALESCE(RTR_PRGRMD_I_H,0)), 0) AS TOTAL_PENSIONADOS_INVALIDEZ
                """ + filtroBase();
    }

    private String sqlVejez() {
        return """
                SELECT COALESCE(
                      SUM(COALESCE(RTR_PRGRMD_SIN_NGCCN_BONO_V_M,0))
                    + SUM(COALESCE(RTR_PRGRMD_SIN_NGCCN_BONO_V_H,0))
                    + SUM(COALESCE(RNT_TMP_VRBL_RNT_VTLC_INMD_V_M,0))
                    + SUM(COALESCE(RNT_TMP_VRBL_RNT_VTLC_INMD_V_H,0))
                    + SUM(COALESCE(RNT_TMP_VRBL_RNT_VTLC_DIF_V_M,0))
                    + SUM(COALESCE(RNT_TMP_VRBL_RNT_VTLC_DIF_V_H,0))
                    + SUM(COALESCE(RNT_VTLC_INMDT_V_M,0))
                    + SUM(COALESCE(RNT_VTLC_INMDT_V_H,0))
                    + SUM(COALESCE(RNT_TMP_CRT_VTLC_DFM_CRT_V_M,0))
                    + SUM(COALESCE(RNT_TMP_CRT_VTLC_DFM_CRT_V_H,0))
                    + SUM(COALESCE(PNSNS_FLL_JDCL_V_M,0))
                    + SUM(COALESCE(PNSNS_FLL_JDCL_V_H,0))
                    + SUM(COALESCE(RTR_PRGRMD_RNT_VTLC_DIF_V_M,0))
                    + SUM(COALESCE(RTR_PRGRMD_RNT_VTLC_DIF_V_H,0))
                    + SUM(COALESCE(RTR_PRGRMD_V_M,0))
                    + SUM(COALESCE(RTR_PRGRMD_V_H,0)), 0) AS TOTAL_PENSIONADOS_VEJEZ
                """ + filtroBase();
    }

    private String sqlSobrevivencia() {
        return """
                SELECT COALESCE(
                      SUM(COALESCE(RNT_TMP_VRBL_RNT_VTLC_INMD_S_M,0))
                    + SUM(COALESCE(RNT_TMP_VRBL_RNT_VTLC_INMD_S_H,0))
                    + SUM(COALESCE(RNT_TMP_VRBL_RNT_VTLC_DIF_S_M,0))
                    + SUM(COALESCE(RNT_TMP_VRBL_RNT_VTLC_DIF_S_H,0))
                    + SUM(COALESCE(RNT_VTLC_INMDT_S_M,0))
                    + SUM(COALESCE(RNT_VTLC_INMDT_S_H,0))
                    + SUM(COALESCE(RNT_TMP_CRT_VTLC_DFM_CRT_S_M,0))
                    + SUM(COALESCE(RNT_TMP_CRT_VTLC_DFM_CRT_S_H,0))
                    + SUM(COALESCE(PNSNS_FLL_JDCL_S_M,0))
                    + SUM(COALESCE(PNSNS_FLL_JDCL_S_H,0))
                    + SUM(COALESCE(RTR_PRGRMD_RNT_VTLC_DIF_S,0))
                    + SUM(COALESCE(RTR_PRGRMD_RNT_VTLC_DIF_S_M,0))
                    + SUM(COALESCE(RTR_PRGRMD_S_M,0))
                    + SUM(COALESCE(RTR_PRGRMD_S_H,0)), 0) AS TOTAL_PENSIONADOS_SOBREVIVIENTES
                """ + filtroBase();
    }

    public record PensionadosResumen(BigDecimal total, BigDecimal invalidez, BigDecimal vejez, BigDecimal sobrevivencia) {}
}
