package co.gov.sfc.excel;

import co.gov.sfc.config.AiosProperties;
import co.gov.sfc.insumos.InsumosLocator;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.lang.reflect.Method;
import java.math.BigDecimal;
import java.nio.file.Path;
import java.time.LocalDate;

import static org.junit.jupiter.api.Assertions.assertEquals;

class SemestralExcelGeneratorTest {

    @Test
    void shouldCreateMissingSemesterColumnAndCopyPreviousFormat() throws Exception {
        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("Hoja1");
            Row months = sheet.createRow(0);
            Row years = sheet.createRow(1);
            months.createCell(2).setCellValue("junio");
            years.createCell(2).setCellValue(2025);
            Row data = sheet.createRow(2);
            Cell source = data.createCell(2);
            CellStyle sourceStyle = workbook.createCellStyle();
            sourceStyle.setBorderRight(BorderStyle.THIN);
            source.setCellStyle(sourceStyle);
            sheet.setColumnWidth(2, 4200);

            SemestralExcelGenerator generator = new SemestralExcelGenerator(
                    new AiosProperties(Path.of("."), Path.of("."), Path.of("."), 40, true), null, null);

            int column = generator.columnaSemestral(sheet, LocalDate.of(2025, 12, 31));

            assertEquals(4, column);
            assertEquals("diciembre", months.getCell(3).getStringCellValue());
            assertEquals(2025, (int) years.getCell(3).getNumericCellValue());
            assertEquals(4200, sheet.getColumnWidth(3));
            assertEquals(BorderStyle.THIN, data.getCell(3).getCellStyle().getBorderRight());
        }
    }

    @Test
    void shouldReadFila25FromFallecidosSheetForJune2025() throws Exception {
        AiosProperties properties = new AiosProperties(Path.of("insumos_ejemplo"), null, null, null, null);
        SemestralExcelGenerator generator = new SemestralExcelGenerator(properties, new InsumosLocator(properties), null);
        Method method = SemestralExcelGenerator.class.getDeclaredMethod("readFila25Trimestral493", LocalDate.class);
        method.setAccessible(true);

        BigDecimal value = (BigDecimal) method.invoke(generator, LocalDate.of(2025, 6, 30));

        assertEquals(new BigDecimal("38.27900000"), value);
    }
}
