package co.gov.sfc.excel;

import co.gov.sfc.config.AiosProperties;
import org.junit.jupiter.api.Test;

import java.math.BigDecimal;
import java.nio.file.Path;
import java.time.LocalDate;
import java.util.Map;

import static org.junit.jupiter.api.Assertions.assertTrue;

class TrimestralExcelGeneratorTest {

    @Test
    void shouldGenerateQuarterlyWorkbookFromReferenceTemplate() {
        TrimestralExcelGenerator generator = new TrimestralExcelGenerator(
                new AiosProperties(Path.of("insumos_ejemplo"), Path.of("plantillas"), Path.of("salidas_referencia"), 40, true)
        );

        TrimestralData data = new TrimestralData(
                "jun-25",
                Map.of("mod_colf", BigDecimal.valueOf(1000), "con_colf", BigDecimal.valueOf(800), "mr_colf", BigDecimal.valueOf(100), "mod_sk_total", BigDecimal.valueOf(500)),
                Map.of("colf", BigDecimal.valueOf(1000), "porv", BigDecimal.valueOf(2000), "prot", BigDecimal.valueOf(3000), "sk", BigDecimal.valueOf(4000)),
                Map.of("colf", BigDecimal.valueOf(10000), "porv", BigDecimal.valueOf(20000), "prot", BigDecimal.valueOf(30000), "sk", BigDecimal.valueOf(40000)),
                Map.of("mod_colf", BigDecimal.valueOf(500)),
                Map.of("colf", BigDecimal.valueOf(1)),
                Map.of("col_obl", BigDecimal.valueOf(3.0)),
                Map.of("colf", BigDecimal.valueOf(10.5)),
                Map.of("colf", BigDecimal.valueOf(5.2))
        );

        Path out = generator.generar(LocalDate.of(2025, 6, 30), data);
        assertTrue(out.toFile().exists());
    }
}
