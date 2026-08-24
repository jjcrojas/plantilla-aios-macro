package co.gov.sfc.excel;

import org.junit.jupiter.api.Test;

import java.math.BigDecimal;
import java.net.URI;
import java.time.Month;
import java.util.Map;

import static org.junit.jupiter.api.Assertions.assertEquals;

class ComisionesSfcServiceTest {

    @Test
    void shouldDiscoverAnnualLettersPageFromSecondTableColumn() {
        String html = """
                <table><tr><th>Circulares</th><th>Cartas Circulares (2)</th><th>Resoluciones</th></tr>
                <tr><td><a href='/10115459'>2025</a></td><td><a href='/10115460'>2025</a></td><td><a href='/10115461'>2025</a></td></tr></table>
                """;

        assertEquals(URI.create("https://www.superfinanciera.gov.co/10115460"),
                ComisionesSfcService.findAnnualLettersPage(html, 2025));
    }

    @Test
    void shouldSelectCommissionLetterBySubjectAndPublicationMonth() {
        String html = """
                <table>
                <tr><td><a href='/loader.php?idFile=1'>24</a></td><td>Abril 21</td><td>Otro asunto</td></tr>
                <tr><td><a href='/loader.php?idFile=1075381'>25</a></td><td>Abril 25</td><td>Pública la rentabilidad, comisión de administración y seguro previsional de los Fondos</td></tr>
                </table>
                """;

        assertEquals(URI.create("https://www.superfinanciera.gov.co/loader.php?idFile=1075381"),
                ComisionesSfcService.findCommissionPdf(html, Month.APRIL));
    }

    @Test
    void shouldParseEightValuesAndRepairOcrCommissionUsingThreePercentInvariant() {
        String ocr = """
                PORVENIR 0.47% 2.53% 1.50% 11.50% 16.00%
                COLFONBOS 0.97% 2.03% 1.50% 11.50% 16.00%
                PROTECCION 9.47% 2.53% 1.50% 11.50% 16.00%
                SKANDIA 20506 0.95% 1.50% 11.50% 16.00%
                OTRAS COMISIONES AUTORIZADAS
                COLFONDOS 4.50% 0.49% DEL ULTIMO IBC
                """;

        Map<String, BigDecimal> values = ComisionesSfcService.parseCommissionTable(ocr);

        assertEquals(new BigDecimal("0.97"), values.get("col_obl"));
        assertEquals(new BigDecimal("2.03"), values.get("col_seg"));
        assertEquals(new BigDecimal("0.47"), values.get("por_obl"));
        assertEquals(new BigDecimal("2.53"), values.get("por_seg"));
        assertEquals(new BigDecimal("0.47"), values.get("pro_obl"));
        assertEquals(new BigDecimal("2.53"), values.get("pro_seg"));
        assertEquals(new BigDecimal("2.05"), values.get("ska_obl"));
        assertEquals(new BigDecimal("0.95"), values.get("ska_seg"));
    }
}
