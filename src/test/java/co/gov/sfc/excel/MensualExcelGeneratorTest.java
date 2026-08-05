package co.gov.sfc.excel;

import co.gov.sfc.config.AiosProperties;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.nio.file.Path;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;

class MensualExcelGeneratorTest {

    @Test
    void shouldCreateMissingMonthlyPeriodAfterLastExistingPeriod() throws Exception {
        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("HOJA1");
            sheet.createRow(0).createCell(0).setCellValue("Fecha");
            Row junio = sheet.createRow(4);
            junio.createCell(0).setCellValue("jun-25");
            junio.createCell(1).setCellValue(100);
            sheet.createRow(5); // fila vacía ya presente en la plantilla
            sheet.createRow(6).createCell(0).setCellValue("(1) Nota metodológica");

            AiosProperties properties = new AiosProperties(Path.of("."), Path.of("."), Path.of("."), 40, false);
            MensualExcelGenerator generator = new MensualExcelGenerator(properties, new CeldaLogger());

            int row = generator.findOrCreateDateRow(sheet, "jul-25");
            generator.aplicarFormatoFilaMensual(sheet, row);

            assertEquals(6, row);
            assertEquals("jul-25", sheet.getRow(5).getCell(0).getStringCellValue());
            assertEquals("#,##0", sheet.getRow(5).getCell(1).getCellStyle().getDataFormatString());
            assertEquals("(1) Nota metodológica", sheet.getRow(7).getCell(0).getStringCellValue());
        }
    }

    @Test
    void shouldApplyMacroFormatUsingNearestPopulatedRowAcrossBlankMonths() throws Exception {
        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("HOJA1");
            Row reference = sheet.createRow(4);
            reference.setHeightInPoints(18);
            Font font = workbook.createFont();
            font.setFontName("Arial");
            font.setBold(false);
            CellStyle referenceStyle = workbook.createCellStyle();
            referenceStyle.setFont(font);
            referenceStyle.setBorderBottom(BorderStyle.THIN);
            for (int col = 0; col < 19; col++) {
                var cell = reference.createCell(col);
                cell.setCellStyle(referenceStyle);
                if (col > 0) cell.setCellValue(col);
            }

            sheet.createRow(5).createCell(0).setCellValue("jun-25");
            sheet.createRow(6).createCell(0).setCellValue("jul-25");
            sheet.createRow(7).createCell(0).setCellValue("ago-25");

            AiosProperties properties = new AiosProperties(Path.of("."), Path.of("."), Path.of("."), 40, false);
            MensualExcelGenerator generator = new MensualExcelGenerator(properties, new CeldaLogger());
            generator.aplicarFormatoFilaMensual(sheet, 8);

            Row formatted = sheet.getRow(7);
            assertEquals(reference.getHeight(), formatted.getHeight());
            assertEquals("#,##0", formatted.getCell(1).getCellStyle().getDataFormatString());
            assertEquals("#,##0.00", formatted.getCell(6).getCellStyle().getDataFormatString());
            assertEquals("#,##0", formatted.getCell(15).getCellStyle().getDataFormatString());
            assertEquals("#,##0.00", formatted.getCell(18).getCellStyle().getDataFormatString());
            assertEquals("Arial", workbook.getFontAt(formatted.getCell(6).getCellStyle().getFontIndex()).getFontName());
            assertFalse(workbook.getFontAt(formatted.getCell(6).getCellStyle().getFontIndex()).getBold());
            assertEquals(BorderStyle.THIN, formatted.getCell(6).getCellStyle().getBorderBottom());
        }
    }
}
