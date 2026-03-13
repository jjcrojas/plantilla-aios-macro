package co.gov.sfc.excel;

import co.gov.sfc.config.AiosProperties;
import org.junit.jupiter.api.Test;

import java.math.BigDecimal;
import java.nio.file.Path;
import java.time.LocalDate;

import static org.junit.jupiter.api.Assertions.assertTrue;

class TrimestralExcelGeneratorTest {

    @Test
    void shouldGenerateQuarterlyWorkbookFromReferenceTemplate() {
        TrimestralExcelGenerator generator = new TrimestralExcelGenerator(
                new AiosProperties(Path.of("insumos_ejemplo"), Path.of("plantillas"), Path.of("salidas_referencia"), 40, true)
        );

        TrimestralData data = new TrimestralData(
                "jun-25",
                BigDecimal.valueOf(1_000_000),
                BigDecimal.valueOf(2_000_000),
                BigDecimal.valueOf(3_000_000),
                BigDecimal.valueOf(4_000_000),
                BigDecimal.valueOf(500_000),
                BigDecimal.valueOf(10_000),
                BigDecimal.valueOf(20_000),
                BigDecimal.valueOf(30_000),
                BigDecimal.valueOf(40_000),
                BigDecimal.valueOf(10.5),
                BigDecimal.valueOf(5.2)
        );

        Path out = generator.generar(LocalDate.of(2025, 6, 30), data);
        assertTrue(out.toFile().exists());
    }
}
