package co.gov.sfc.excel;

import co.gov.sfc.config.AiosProperties;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.springframework.context.annotation.AnnotationConfigApplicationContext;

import java.math.BigDecimal;
import java.nio.file.Path;
import java.time.LocalDate;
import java.util.Map;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

class TrimestralExcelGeneratorTest {

    @Test
    void shouldBeCreatedBySpringWhenTestConstructorAlsoExists() {
        AiosProperties properties = new AiosProperties(
                Path.of("target", "insumos-inexistentes"),
                Path.of("target", "plantillas-inexistentes"),
                Path.of("target", "referencias-inexistentes"),
                40,
                true);

        try (AnnotationConfigApplicationContext context = new AnnotationConfigApplicationContext()) {
            context.registerBean(AiosProperties.class, () -> properties);
            context.register(TrimestralExcelGenerator.class);
            context.refresh();

            assertTrue(context.getBean(TrimestralExcelGenerator.class) != null);
        }
    }

    @Test
    void shouldCreateMissingQuarterAndCopyPreviousRowFormat() throws Exception {
        try (Workbook workbook = new XSSFWorkbook()) {
            var sheet = workbook.createSheet("afiliados");
            var junio = sheet.createRow(9);
            junio.setHeightInPoints(17);
            junio.createCell(0).setCellValue("jun-25");
            var sourceCell = junio.createCell(1);
            var sourceStyle = workbook.createCellStyle();
            sourceStyle.setBorderBottom(BorderStyle.DOTTED);
            sourceCell.setCellStyle(sourceStyle);
            sheet.createRow(10); // fila vacía existente
            sheet.createRow(11).createCell(0).setCellValue("(1) Incluye fondo alternativo");

            TrimestralExcelGenerator generator = new TrimestralExcelGenerator(
                    new AiosProperties(Path.of("."), Path.of("."), Path.of("."), 40, true));

            int row = generator.findOrAppendRow(sheet, LocalDate.of(2025, 9, 30), "sep-25");

            assertEquals(11, row);
            assertEquals("sep-25", sheet.getRow(10).getCell(0).getStringCellValue());
            assertEquals(junio.getHeight(), sheet.getRow(10).getHeight());
            assertEquals(BorderStyle.DOTTED, sheet.getRow(10).getCell(1).getCellStyle().getBorderBottom());
            assertEquals("(1) Incluye fondo alternativo", sheet.getRow(11).getCell(0).getStringCellValue());
        }
    }

    @Test
    void shouldGenerateQuarterlyWorkbookFromBlankInternalTemplate() throws Exception {
        Path missingRoot = Path.of("target", "plantillas-inexistentes-prueba");
        TrimestralExcelGenerator generator = new TrimestralExcelGenerator(
                new AiosProperties(missingRoot.resolve("insumos"), missingRoot.resolve("plantillas"),
                        missingRoot.resolve("salidas_referencia"), 40, true)
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
        try (Workbook workbook = org.apache.poi.ss.usermodel.WorkbookFactory.create(out.toFile())) {
            assertEquals("jun-25", workbook.getSheet("afiliados").getRow(7).getCell(0).getStringCellValue());
            assertEquals(1000d, workbook.getSheet("afiliados").getRow(7).getCell(1).getNumericCellValue());
            assertEquals("jun-25", workbook.getSheet("aportantes").getRow(6).getCell(0).getStringCellValue());
            assertEquals("jun-25", workbook.getSheet("gastos").getRow(13).getCell(0).getStringCellValue());
            assertTrue(workbook.getSheet("afiliados").getRow(7).getCell(1).getCellStyle().getIndex() != 0);
        }
    }

    @Test
    void shouldSortExistingPeriodsChronologicallyAndNormalizeSeptemberLabel() throws Exception {
        try (Workbook workbook = new XSSFWorkbook()) {
            var sheet = workbook.createSheet("afiliados");
            sheet.createRow(14).createCell(0).setCellValue("jun-25");
            sheet.createRow(15).createCell(0).setCellValue("dic-25");
            sheet.createRow(16).createCell(0).setCellValue("sept-25");

            TrimestralExcelGenerator generator = new TrimestralExcelGenerator(
                    new AiosProperties(Path.of("."), Path.of("."), Path.of("."), 40, true));

            int row = generator.findOrAppendRow(sheet, LocalDate.of(2025, 9, 30), "sept-25");

            assertEquals(16, row);
            assertEquals("jun-25", sheet.getRow(14).getCell(0).getStringCellValue());
            assertEquals("sep-25", sheet.getRow(15).getCell(0).getStringCellValue());
            assertEquals("dic-25", sheet.getRow(16).getCell(0).getStringCellValue());
        }
    }
}
