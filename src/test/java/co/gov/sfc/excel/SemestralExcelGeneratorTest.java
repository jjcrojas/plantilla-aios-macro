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
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.LocalDate;
import java.util.Map;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;
import static org.mockito.ArgumentMatchers.any;
import static org.mockito.ArgumentMatchers.anyInt;
import static org.mockito.Mockito.mock;
import static org.mockito.Mockito.verify;
import static org.mockito.Mockito.when;

class SemestralExcelGeneratorTest {

    @Test
    void shouldGenerateFromBlankInternalTemplateWithoutReferenceDirectory() throws Exception {
        LocalDate cutoff = LocalDate.of(2025, 6, 30);
        Path missingRoot = Path.of("target", "sin-referencias-semestral");
        AiosProperties properties = new AiosProperties(
                missingRoot.resolve("insumos"), missingRoot.resolve("plantillas"),
                missingRoot.resolve("salidas_referencia"), 40, true);
        InsumosLocator locator = mock(InsumosLocator.class);
        RentabilidadService rentabilidadService = mock(RentabilidadService.class);
        Formato493QueryService formato493 = mock(Formato493QueryService.class);
        Formato495QueryService formato495 = mock(Formato495QueryService.class);
        Formato136QueryService formato136 = mock(Formato136QueryService.class);
        ComisionesSemestralQueryService comisiones = mock(ComisionesSemestralQueryService.class);
        Path rentFile = missingRoot.resolve("Rent_Vr_Uni_Moderado.xlsm");

        when(locator.findRequired("Rent_Vr_Uni_Moderado", cutoff)).thenReturn(rentFile);
        when(rentabilidadService.calcularRentabilidad(any(Path.class), any(LocalDate.class), anyInt()))
                .thenAnswer(invocation -> {
                    LocalDate end = invocation.getArgument(1);
                    int years = invocation.getArgument(2);
                    return new RentabilidadService.RentabilidadResultado(
                            end.minusYears(years), end, BigDecimal.ZERO, BigDecimal.ZERO);
                });
        when(formato493.leerFallecidosSistema(cutoff)).thenReturn(BigDecimal.ZERO);
        when(formato495.leerResumen(cutoff)).thenReturn(new Formato495QueryService.PensionadosResumen(
                BigDecimal.ZERO, BigDecimal.ZERO, BigDecimal.ZERO, BigDecimal.ZERO));
        when(formato136.leerAportesRecibidos(cutoff)).thenReturn(BigDecimal.ZERO);
        when(comisiones.leer590000(any(), any())).thenReturn(BigDecimal.ZERO);
        when(comisiones.leer411500(any(), any())).thenReturn(BigDecimal.ZERO);
        when(comisiones.leerGastosOperativos(any(), any())).thenReturn(BigDecimal.ZERO);
        when(comisiones.leerCuenta(any(), any(), anyInt())).thenReturn(BigDecimal.ZERO);

        SemestralExcelGenerator generator = new SemestralExcelGenerator(
                properties, locator, rentabilidadService, formato493, formato495, formato136, comisiones);
        TrimestralData quarterly = new TrimestralData(
                "jun-25", Map.of(), Map.of(), Map.of(), Map.of(), Map.of(),
                Map.of("col_obl", BigDecimal.ZERO, "por_obl", BigDecimal.ZERO,
                        "pro_obl", BigDecimal.ZERO, "ska_obl", BigDecimal.ZERO),
                Map.of(), Map.of());

        Path output = generator.generar(cutoff, monthlyData("jun-25"), quarterly);

        assertTrue(Files.isRegularFile(output));
        try (Workbook workbook = org.apache.poi.ss.usermodel.WorkbookFactory.create(output.toFile())) {
            Sheet sheet = workbook.getSheet("Hoja1");
            assertEquals("junio", sheet.getRow(0).getCell(2).getStringCellValue());
            assertEquals(2025, (int) sheet.getRow(1).getCell(2).getNumericCellValue());
            assertEquals(1d, sheet.getRow(2).getCell(2).getNumericCellValue());
            assertEquals(0d, sheet.getRow(41).getCell(2).getNumericCellValue());
            assertNotEquals(0, sheet.getRow(2).getCell(2).getCellStyle().getIndex());
        }
    }

    @Test
    void shouldUseFirstFormattedColumnOfBlankInternalTemplate() throws Exception {
        Path missingRoot = Path.of("target", "plantillas-inexistentes-semestral");
        AiosProperties properties = new AiosProperties(
                missingRoot.resolve("insumos"), missingRoot.resolve("plantillas"),
                missingRoot.resolve("salidas_referencia"), 40, true);

        try (Workbook workbook = new AiosTemplateService(properties).openWorkbook("semestral.xlsx")) {
            Sheet sheet = workbook.getSheet("Hoja1");
            short originalStyle = sheet.getRow(2).getCell(2).getCellStyle().getIndex();
            SemestralExcelGenerator generator = new SemestralExcelGenerator(
                    properties, null, null, null, null, null, null);

            int column = generator.columnaSemestral(sheet, LocalDate.of(2025, 6, 30));

            assertEquals(3, column);
            assertEquals("junio", sheet.getRow(0).getCell(2).getStringCellValue());
            assertEquals(2025, (int) sheet.getRow(1).getCell(2).getNumericCellValue());
            assertEquals(originalStyle, sheet.getRow(2).getCell(2).getCellStyle().getIndex());
        }
    }

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
        AiosProperties properties = new AiosProperties(Path.of("target", "insumos-dinamicos-prueba"), null, null, null, null);
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

    @Test
    void shouldKeepJuneAndDecemberInChronologicalColumns() throws Exception {
        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("Hoja1");
            sheet.createRow(0);
            sheet.createRow(1);
            sheet.createRow(2).createCell(1).setCellValue(1);

            SemestralExcelGenerator generator = new SemestralExcelGenerator(
                    new AiosProperties(Path.of("."), Path.of("."), Path.of("."), 40, true),
                    null, null, null, null, null, null);

            int junio = generator.columnaSemestral(sheet, LocalDate.of(2025, 6, 30));
            int diciembre = generator.columnaSemestral(sheet, LocalDate.of(2025, 12, 31));

            assertEquals(3, junio);
            assertEquals(4, diciembre);
            assertEquals("junio", sheet.getRow(0).getCell(2).getStringCellValue());
            assertEquals("diciembre", sheet.getRow(0).getCell(3).getStringCellValue());
            assertEquals(2025, (int) sheet.getRow(1).getCell(2).getNumericCellValue());
            assertEquals(2025, (int) sheet.getRow(1).getCell(3).getNumericCellValue());
        }
    }

    @Test
    void shouldWriteRow47AsNumericPercentageWithoutPercentSymbol() throws Exception {
        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("Hoja1");
            SemestralExcelGenerator generator = new SemestralExcelGenerator(
                    new AiosProperties(Path.of("."), Path.of("."), Path.of("."), 40, true),
                    null, null, null, null, null, null);

            generator.writeFila47SinSimboloPorcentaje(sheet, 3, new BigDecimal("70.25"));

            Cell cell = sheet.getRow(46).getCell(2);
            assertEquals(70.25d, cell.getNumericCellValue());
            assertEquals("#,##0.00", cell.getCellStyle().getDataFormatString());
        }
    }

    @Test
    void shouldWriteRow27AsNumericPercentageWithoutPercentSymbol() throws Exception {
        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("Hoja1");
            SemestralExcelGenerator generator = new SemestralExcelGenerator(
                    new AiosProperties(Path.of("."), Path.of("."), Path.of("."), 40, true),
                    null, null, null, null, null, null);

            generator.writeFila27SinSimboloPorcentaje(sheet, 3, new BigDecimal("0.62"));

            Cell cell = sheet.getRow(26).getCell(2);
            assertEquals(0.62d, cell.getNumericCellValue());
            assertEquals("#,##0.00", cell.getCellStyle().getDataFormatString());
        }
    }

    @Test
    void shouldWriteRow26WithoutDecimals() throws Exception {
        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("Hoja1");
            SemestralExcelGenerator generator = new SemestralExcelGenerator(
                    new AiosProperties(Path.of("."), Path.of("."), Path.of("."), 40, true),
                    null, null, null, null, null, null);

            generator.writeFila26SinDecimales(sheet, 3, new BigDecimal("123456"));

            Cell cell = sheet.getRow(25).getCell(2);
            assertEquals(123456d, cell.getNumericCellValue());
            assertEquals("#,##0", cell.getCellStyle().getDataFormatString());
        }
    }

    @Test
    void shouldCalculateRow60AsTheSumOfRows55_56_57And59() {
        LocalDate fechaCorte = LocalDate.of(2025, 6, 30);
        BigDecimal trm = new BigDecimal("4069.67");
        ComisionesSemestralQueryService queryService = mock(ComisionesSemestralQueryService.class);
        when(queryService.leerCuenta(fechaCorte, trm, 512000)).thenReturn(new BigDecimal("100"));
        when(queryService.leerCuenta(fechaCorte, trm, 513000)).thenReturn(new BigDecimal("50"));
        when(queryService.leerCuenta(fechaCorte, trm, 511524)).thenReturn(new BigDecimal("60"));
        when(queryService.leerCuenta(fechaCorte, trm, 511527)).thenReturn(new BigDecimal("40"));
        when(queryService.leerCuenta(fechaCorte, trm, 519015)).thenReturn(new BigDecimal("30"));
        SemestralExcelGenerator generator = new SemestralExcelGenerator(
                new AiosProperties(Path.of("."), Path.of("."), Path.of("."), 40, true),
                null, null, null, null, null, queryService);

        SemestralExcelGenerator.FilasGastosSemestrales filas = generator.calcularFilasGastosSemestrales(
                fechaCorte, trm, new BigDecimal("1000"), new BigDecimal("700"));

        assertEquals(new BigDecimal("300"), filas.fila53());
        assertEquals(new BigDecimal("150"), filas.fila55());
        assertEquals(new BigDecimal("100"), filas.fila56());
        assertEquals(new BigDecimal("30"), filas.fila57());
        assertEquals(new BigDecimal("130"), filas.fila58());
        assertEquals(new BigDecimal("420"), filas.fila59());
        assertEquals(new BigDecimal("700"), filas.fila60());
        verify(queryService).leerCuenta(fechaCorte, trm, 512000);
        verify(queryService).leerCuenta(fechaCorte, trm, 513000);
        verify(queryService).leerCuenta(fechaCorte, trm, 511524);
        verify(queryService).leerCuenta(fechaCorte, trm, 511527);
        verify(queryService).leerCuenta(fechaCorte, trm, 519015);
    }

    @Test
    void shouldCalculateRows61_65And66FromTheReferencedRows() {
        SemestralExcelGenerator generator = new SemestralExcelGenerator(
                new AiosProperties(Path.of("."), Path.of("."), Path.of("."), 40, true),
                null, null, null, null, null, null);

        SemestralExcelGenerator.IndicadoresFinancierosSemestrales indicadores =
                generator.calcularIndicadoresFinancierosSemestrales(
                        new BigDecimal("42000000000000"),
                        new BigDecimal("7500000"),
                        new BigDecimal("4000"),
                        new BigDecimal("150"),
                        new BigDecimal("750"),
                        new BigDecimal("1500"));

        assertEquals(new BigDecimal("1400.00000000"), indicadores.fila61());
        assertEquals(new BigDecimal("20.00000000"), indicadores.fila65());
        assertEquals(new BigDecimal("10.00000000"), indicadores.fila66());
    }

    private MensualData monthlyData(String period) {
        BigDecimal one = BigDecimal.ONE;
        return new MensualData(period,
                one, one, one, one, one, one,
                one, one, one, one, one, one, one,
                one, one, one, one, one, one,
                one, one, one, one, one, one,
                one, one, one, one, one, one,
                one, one, one, one, one, one,
                one, one, one, one, one,
                one, one, one);
    }
}
