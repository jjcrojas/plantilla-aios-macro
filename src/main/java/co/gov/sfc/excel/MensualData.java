package co.gov.sfc.excel;

import java.math.BigDecimal;

public record MensualData(
        String textoFecha,
        BigDecimal afiliados,
        BigDecimal aportantes,
        BigDecimal traspasosSistema,
        BigDecimal vrFondo,
        BigDecimal trm,
        BigDecimal tmpNominal1,
        BigDecimal tmpReal1,
        BigDecimal consFdosAdmon,
        BigDecimal porcVrFondo,
        BigDecimal total1,
        BigDecimal dudaG,
        BigDecimal dudaEf,
        BigDecimal dudaNf,
        BigDecimal dudaAc,
        BigDecimal dudaF,
        BigDecimal h17,
        BigDecimal otros
) {
}
