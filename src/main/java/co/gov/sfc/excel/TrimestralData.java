package co.gov.sfc.excel;

import java.math.BigDecimal;

public record TrimestralData(
        String etiquetaFecha,
        BigDecimal cotColfondos,
        BigDecimal cotPorvenir,
        BigDecimal cotProteccion,
        BigDecimal cotSkandia,
        BigDecimal vrFondoUsd,
        BigDecimal traspasosSistema,
        BigDecimal rentNominal12m,
        BigDecimal rentReal12m
) {
}

