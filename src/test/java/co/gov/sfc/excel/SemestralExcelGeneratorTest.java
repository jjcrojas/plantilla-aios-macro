package co.gov.sfc.excel;

import co.gov.sfc.config.AiosProperties;
import co.gov.sfc.insumos.InsumosLocator;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.lang.reflect.Method;
import java.math.BigDecimal;
import java.nio.file.Path;
import java.time.LocalDate;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.mockito.Mockito.mock;
import static org.mockito.Mockito.when;

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
                    new AiosProperties(Path.of("."), Path.of("."), Path.of("."), 40, true),
                    null, null, null, null, null, null);

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
        Formato493QueryService formato493QueryService = mock(Formato493QueryService.class);
        when(formato493QueryService.leerFallecidosSistema(LocalDate.of(2025, 6, 30))).thenReturn(new BigDecimal("38279"));
        SemestralExcelGenerator generator = new SemestralExcelGenerator(properties, new InsumosLocator(properties), null, formato493QueryService, mock(Formato495QueryService.class), mock(Formato136QueryService.class), null);
        Method method = SemestralExcelGenerator.class.getDeclaredMethod("readFila25Trimestral493", LocalDate.class);
        method.setAccessible(true);

        BigDecimal value = (BigDecimal) method.invoke(generator, LocalDate.of(2025, 6, 30));

        assertEquals(new BigDecimal("38.27900000"), value);
    }

    @Test
    void shouldUseBlackFontsAndRemoveGreenOrangeBand() throws Exception {
        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("Hoja1");
            Cell redCell = sheet.createRow(68).createCell(2);
            Font redFont = workbook.createFont();
            redFont.setColor(IndexedColors.RED.getIndex());
            CellStyle redStyle = workbook.createCellStyle();
            redStyle.setFont(redFont);
            redCell.setCellStyle(redStyle);

            Row band = sheet.createRow(80);
            Cell green = band.createCell(2);
            CellStyle greenStyle = workbook.createCellStyle();
            greenStyle.setFillForegroundColor(IndexedColors.BRIGHT_GREEN.getIndex());
            greenStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            green.setCellStyle(greenStyle);
            Cell orange = band.createCell(3);
            CellStyle orangeStyle = workbook.createCellStyle();
            orangeStyle.setFillForegroundColor(IndexedColors.ORANGE.getIndex());
            orangeStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            orange.setCellStyle(orangeStyle);

            SemestralExcelGenerator generator = new SemestralExcelGenerator(
                    new AiosProperties(Path.of("."), Path.of("."), Path.of("."), 40, true),
                    null, null, null, null, null, null);
            generator.normalizarEstilosSemestral(sheet);

            assertEquals(IndexedColors.BLACK.getIndex(), workbook.getFontAt(redCell.getCellStyle().getFontIndex()).getColor());
            assertEquals(FillPatternType.NO_FILL, green.getCellStyle().getFillPattern());
            assertEquals(FillPatternType.NO_FILL, orange.getCellStyle().getFillPattern());
        }
    }

    @Test
    void shouldWriteTotalAfiliadosInRow3AndNoDisponibleInRow20() throws Exception {
        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("Hoja1");
            MensualData mensual = mock(MensualData.class);
            when(mensual.afiliados()).thenReturn(new BigDecimal("123456"));

            SemestralExcelGenerator generator = new SemestralExcelGenerator(
                    new AiosProperties(Path.of("."), Path.of("."), Path.of("."), 40, true),
                    null, null, null, null, null, null);
            generator.writeFilasAfiliadosDisponibilidad(sheet, 3, mensual);

            assertEquals(123456d, sheet.getRow(2).getCell(2).getNumericCellValue());
            assertEquals("No Disponible", sheet.getRow(19).getCell(2).getStringCellValue());
        }
    }
}
