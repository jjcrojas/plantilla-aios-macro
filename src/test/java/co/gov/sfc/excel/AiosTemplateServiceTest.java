package co.gov.sfc.excel;

import co.gov.sfc.config.AiosProperties;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.nio.file.Path;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotEquals;

class AiosTemplateServiceTest {

    @TempDir
    Path tempDir;

    @Test
    void shouldLoadBlankFormattedTemplatesWithoutExternalReferenceFiles() throws Exception {
        AiosProperties properties = new AiosProperties(
                tempDir.resolve("insumos"),
                tempDir.resolve("plantillas"),
                tempDir.resolve("salidas_referencia"),
                40,
                false);
        AiosTemplateService templates = new AiosTemplateService(properties);
        DataFormatter formatter = new DataFormatter();

        try (Workbook monthly = templates.openWorkbook("Boletin_AIOS MENSUAL.xlsx")) {
            var sheet = monthly.getSheet("Hoja1");
            assertEquals("Fecha", formatter.formatCellValue(sheet.getRow(0).getCell(0)));
            assertEquals("", formatter.formatCellValue(sheet.getRow(1).getCell(0)));
            assertEquals("", formatter.formatCellValue(sheet.getRow(1).getCell(1)));
            assertNotEquals(0, sheet.getRow(1).getCell(1).getCellStyle().getIndex());
        }

        try (Workbook quarterly = templates.openWorkbook("Boletin_AIOS TRIMESTRAL.xlsx")) {
            assertEquals("", formatter.formatCellValue(quarterly.getSheet("afiliados").getRow(7).getCell(0)));
            assertEquals("", formatter.formatCellValue(quarterly.getSheet("gastos").getRow(13).getCell(0)));
            assertNotEquals(0, quarterly.getSheet("afiliados").getRow(7).getCell(1).getCellStyle().getIndex());
            assertEquals("Dato de promotores no está disponible",
                    formatter.formatCellValue(quarterly.getSheet("promotores").getRow(15).getCell(0)));
        }

        try (Workbook semiannual = templates.openWorkbook("semestral.xlsx")) {
            var sheet = semiannual.getSheet("Hoja1");
            assertEquals("Afiliados activos", formatter.formatCellValue(sheet.getRow(2).getCell(1)));
            assertEquals("", formatter.formatCellValue(sheet.getRow(0).getCell(2)));
            assertEquals("", formatter.formatCellValue(sheet.getRow(2).getCell(2)));
            assertNotEquals(0, sheet.getRow(2).getCell(2).getCellStyle().getIndex());
        }
    }
}
