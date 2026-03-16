package co.gov.sfc.excel;

import java.math.BigDecimal;
import java.util.Map;

public record TrimestralData(
        String etiquetaFecha,
        Map<String, BigDecimal> afiliados,
        Map<String, BigDecimal> aportantes,
        Map<String, BigDecimal> traspasos,
        Map<String, BigDecimal> colombiaUsd,
        Map<String, BigDecimal> gastosUsd,
        Map<String, BigDecimal> comisionesPct,
        Map<String, BigDecimal> rentNominalPct,
        Map<String, BigDecimal> rentRealPct
) {
}
