package co.gov.sfc.excel;

import co.gov.sfc.config.AiosProperties;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellReference;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Component;

import java.io.OutputStream;
import java.math.BigDecimal;
import java.math.RoundingMode;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.LocalDate;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.HashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;

@Component
public class MensualExcelGenerator {

    private final CeldaLogger celdaLogger;
    private final AiosTemplateService templateService;
    private final Path outputDir;

    @Autowired
    public MensualExcelGenerator(AiosProperties properties, CeldaLogger celdaLogger) {
        this(properties, celdaLogger, new AiosTemplateService(properties), Path.of("target", "aios-output"));
    }

    MensualExcelGenerator(AiosProperties properties, CeldaLogger celdaLogger, Path outputDir) {
        this(properties, celdaLogger, new AiosTemplateService(properties), outputDir);
    }

    MensualExcelGenerator(AiosProperties properties, CeldaLogger celdaLogger,
                          AiosTemplateService templateService, Path outputDir) {
        this.celdaLogger = celdaLogger;
        this.templateService = templateService;
        this.outputDir = outputDir;
    }

    public Path generar(MensualData data) {
        return generar(List.of(data));
    }

    public Path generar(List<MensualData> datos) {
        if (datos == null || datos.isEmpty()) {
            throw new IllegalArgumentException("Debe suministrar al menos un período mensual");
        }
        try {
            Files.createDirectories(outputDir);
            Path out = Files.createTempDirectory(outputDir, "mensual-")
                    .resolve("Boletin_AIOS MENSUAL.xlsx");
            try (Workbook wb = templateService.openWorkbook("Boletin_AIOS MENSUAL.xlsx")) {
                Sheet sheet = wb.getSheet("HOJA1");
                if (sheet == null) {
                    throw new IllegalStateException("La plantilla mensual no contiene la hoja HOJA1");
                }
                sortAndNormalizePeriodRows(sheet);
                List<MensualData> datosOrdenados = datos.stream()
                        .sorted(Comparator.comparing(data -> periodDate(data.textoFecha())))
                        .toList();
                for (MensualData data : datosOrdenados) {
                    escribirPeriodo(sheet, data);
                }

                try (OutputStream os = Files.newOutputStream(out)) {
                    wb.write(os);
                }
            }
            return out;
        } catch (Exception e) {
            throw new IllegalStateException("No fue posible generar boletín mensual: " + e.getMessage(), e);
        }
    }

    private void escribirPeriodo(Sheet sheet, MensualData data) {
        int row = findOrCreateDateRow(sheet, data.textoFecha());
        aplicarFormatoFilaMensual(sheet, row);
        write(sheet, row, 2, data.afiliados(), "Query Teradata PROD_DWH_CONSULTA.FORMATO491", "FORMATO491", "SUM(TOTAL_AFILIADOS_TOTAL), RENGLON=999");
                write(sheet, row, 3, data.aportantes(), "Query Teradata PROD_DWH_CONSULTA.FORMATO491", "FORMATO491", "SUM(TOTAL_AFILIADOS_COTIZANTES), RENGLON=999");
                write(sheet, row, 4, data.traspasosSistema(), "Query Teradata PROD_DWH_CONSULTA.S9_FORMATO_493", "S9_FORMATO_493", "Traspasos sistema 12 meses por FECHA_CORTE/UNIDAD_CAPTURA/RENGLON");
                write(sheet, row, 5, divide(data.vrFondo(), trm(data)), "Query Teradata PROD_DWH_CONSULTA.NEGFID_INSUMO_ENTIDAD", "niveles 136/2/4/305", "SUM(valor)/1000000/TRM");
                write(sheet, row, 6, divide(data.total1(), trm(data)), "LIMITES del nuevo.xlsm", "AIOS", "AB4");
                write(sheet, row, 7, pct(data.dudaG()), "LIMITES del nuevo.xlsm", "AIOS", "C4");
                write(sheet, row, 8, pct(data.dudaEf()), "LIMITES del nuevo.xlsm", "AIOS", "E4");
                write(sheet, row, 9, pct(data.dudaNf()), "LIMITES del nuevo.xlsm", "AIOS", "G4");
                write(sheet, row, 10, pct(data.dudaAc()), "LIMITES del nuevo.xlsm", "AIOS", "I4");
                write(sheet, row, 11, pct(data.dudaF()), "LIMITES del nuevo.xlsm", "AIOS", "K4");
                write(sheet, row, 12, pct(data.h17()), "LIMITES del nuevo.xlsm", "AIOS", "O4:Y4");
                write(sheet, row, 13, pct(data.otros()), "LIMITES del nuevo.xlsm", "AIOS", "AA4");
                write(sheet, row, 14, pct(data.tmpNominal1()), "Rent_Vr_Uni_Moderado.xlsm",
                        "Consolidado", "NAV columna E, horizonte exacto de 1 año; cálculo en Java");
                write(sheet, row, 15, pct(data.tmpReal1()), "Rent_Vr_Uni_Moderado.xlsm",
                        "Consolidado + IPC_D", "NAV columna E ajustado por IPC_D columna B; cálculo en Java");
                write(sheet, row, 16,
                        data.administradorasVigentes() == null ? BigDecimal.ZERO : data.administradorasVigentes(),
                        "Query Teradata PROD_DWH_CONSULTA.ENTIDADES", "ENTIDADES",
                        "COUNT nombres distintos, Tipo_Entidad=23, Estado=1; nombres sin comillas");
                write(sheet, row, 17, data.consFdosAdmon(), "Query Teradata PROD_DWH_CONSULTA.FORMATO491", "FORMATO491", "Top 2 AFP por TOTAL_AFILIADOS_TOTAL / total sistema");
                write(sheet, row, 18, data.porcVrFondo(), "Query Teradata PROD_DWH_CONSULTA.NEGFID_INSUMO_ENTIDAD", "niveles 136/2/4/305", "(Proteccion+Porvenir)/total sistema * 100");
                write(sheet, row, 19, trm(data), "Servicio web TRM Superfinanciera (archivo PIB_PEA_TRM_DG como contingencia)", "queryTCRM", "fecha de corte");
    }

    private BigDecimal trm(MensualData data) {
        return data.trm().signum() == 0 ? BigDecimal.ONE : data.trm();
    }

    private BigDecimal divide(BigDecimal a, BigDecimal b) {
        if (b.signum() == 0) return BigDecimal.ZERO;
        return a.divide(b, 8, RoundingMode.HALF_UP);
    }

    private BigDecimal pct(BigDecimal value) {
        return value.multiply(BigDecimal.valueOf(100));
    }

    void aplicarFormatoFilaMensual(Sheet sheet, int row1Based) {
        Row target = sheet.getRow(row1Based - 1);
        if (target == null) target = sheet.createRow(row1Based - 1);

        Row reference = findNearestPopulatedRow(sheet, row1Based);
        if (reference != null && reference.getHeight() > 0) {
            target.setHeight(reference.getHeight());
        }

        DataFormat dataFormat = sheet.getWorkbook().createDataFormat();
        for (int col1Based = 1; col1Based <= 19; col1Based++) {
            Cell targetCell = target.getCell(col1Based - 1);
            if (targetCell == null) targetCell = target.createCell(col1Based - 1);

            CellStyle style = sheet.getWorkbook().createCellStyle();
            Cell referenceCell = reference == null ? null : reference.getCell(col1Based - 1);
            if (referenceCell != null) {
                style.cloneStyleFrom(referenceCell.getCellStyle());
            } else if (targetCell.getCellStyle() != null) {
                style.cloneStyleFrom(targetCell.getCellStyle());
            }
            style.setDataFormat(dataFormat.getFormat(numberFormatForColumn(col1Based)));
            targetCell.setCellStyle(style);
        }
    }

    private Row findNearestPopulatedRow(Sheet sheet, int targetRow1Based) {
        for (int rowIndex = targetRow1Based - 2; rowIndex >= 1; rowIndex--) {
            Row candidate = sheet.getRow(rowIndex);
            if (hasMonthlyData(candidate)) return candidate;
        }
        for (int rowIndex = targetRow1Based; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
            Row candidate = sheet.getRow(rowIndex);
            if (hasMonthlyData(candidate)) return candidate;
        }
        return null;
    }

    private boolean hasMonthlyData(Row row) {
        if (row == null) return false;
        for (int colIndex = 1; colIndex < 19; colIndex++) {
            Cell cell = row.getCell(colIndex);
            if (cell != null && cell.getCellType() != CellType.BLANK) return true;
        }
        return false;
    }

    private String numberFormatForColumn(int col1Based) {
        if (col1Based == 1) return "mmm-yy";
        if (col1Based >= 2 && col1Based <= 6) return "#,##0";
        if (col1Based >= 7 && col1Based <= 15) return "#,##0.00";
        if (col1Based == 16) return "#,##0";
        return "#,##0.00";
    }

    private void write(Sheet sheet, int row1Based, int col1Based, BigDecimal value, String fuenteArchivo, String fuenteHoja, String fuenteCelda) {
        Row row = sheet.getRow(row1Based - 1);
        if (row == null) row = sheet.createRow(row1Based - 1);
        Cell cell = row.getCell(col1Based - 1);
        if (cell == null) cell = row.createCell(col1Based - 1, CellType.NUMERIC);
        cell.setCellValue(value.doubleValue());
        String celda = CellReference.convertNumToColString(col1Based - 1) + row1Based;
        celdaLogger.log(sheet.getSheetName(), celda, value, fuenteArchivo, fuenteHoja, fuenteCelda);
    }

    int findOrCreateDateRow(Sheet sheet, String textoFecha) {
        sortAndNormalizePeriodRows(sheet);
        DataFormatter formatter = new DataFormatter();
        LocalDate fechaObjetivo = periodDate(textoFecha);
        for (Row row : sheet) {
            LocalDate periodo = periodDate(row.getCell(0), formatter);
            if (samePeriod(periodo, fechaObjetivo)) {
                Cell dateCell = row.getCell(0);
                dateCell.setCellValue(canonicalPeriodLabel(fechaObjetivo));
                return row.getRowNum() + 1;
            }
        }

        int lastPeriodRow = -1;
        int insertionRow = -1;
        for (int rowIndex = 1; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
            Row candidate = sheet.getRow(rowIndex);
            if (candidate == null) continue;
            LocalDate periodo = periodDate(candidate.getCell(0), formatter);
            if (periodo == null) continue;
            lastPeriodRow = rowIndex;
            if (insertionRow < 0 && periodo.isAfter(fechaObjetivo)) insertionRow = rowIndex;
        }
        int rowIndex = insertionRow >= 0 ? insertionRow : Math.max(lastPeriodRow + 1, 1);
        Row reusableRow = sheet.getRow(rowIndex);
        if (rowIndex <= sheet.getLastRowNum() && !isReusableBlankMonthlyRow(reusableRow)) {
            sheet.shiftRows(rowIndex, sheet.getLastRowNum(), 1, true, false);
        }
        Row row = sheet.getRow(rowIndex);
        if (row == null) row = sheet.createRow(rowIndex);
        Cell dateCell = row.getCell(0);
        if (dateCell == null) dateCell = row.createCell(0, CellType.STRING);
        dateCell.setCellValue(canonicalPeriodLabel(fechaObjetivo));
        return rowIndex + 1;
    }

    private boolean isReusableBlankMonthlyRow(Row row) {
        if (row == null) return false;
        DataFormatter formatter = new DataFormatter();
        for (int column = 0; column < 19; column++) {
            Cell cell = row.getCell(column);
            if (cell != null && !formatter.formatCellValue(cell).isBlank()) return false;
        }
        return true;
    }

    private void sortAndNormalizePeriodRows(Sheet sheet) {
        DataFormatter formatter = new DataFormatter();
        List<PeriodRowSnapshot> periods = new ArrayList<>();
        for (int rowIndex = 1; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
            Row row = sheet.getRow(rowIndex);
            if (row == null) continue;
            LocalDate period = periodDate(row.getCell(0), formatter);
            if (period != null) periods.add(new PeriodRowSnapshot(rowIndex, period, snapshot(row)));
        }
        if (periods.isEmpty()) return;

        List<PeriodRowSnapshot> ordered = periods.stream()
                .sorted(Comparator.comparing(PeriodRowSnapshot::period))
                .toList();
        for (int i = 0; i < periods.size(); i++) {
            Row target = sheet.getRow(periods.get(i).rowIndex());
            if (target == null) target = sheet.createRow(periods.get(i).rowIndex());
            restore(target, ordered.get(i).row());
            Cell dateCell = target.getCell(0);
            if (dateCell != null) dateCell.setCellValue(canonicalPeriodLabel(ordered.get(i).period()));
        }
    }

    private RowSnapshot snapshot(Row row) {
        Map<Integer, CellSnapshot> cells = new HashMap<>();
        for (Cell cell : row) {
            Object value = switch (cell.getCellType()) {
                case STRING -> cell.getStringCellValue();
                case NUMERIC -> cell.getNumericCellValue();
                case BOOLEAN -> cell.getBooleanCellValue();
                case FORMULA -> cell.getCellFormula();
                case ERROR -> cell.getErrorCellValue();
                default -> null;
            };
            cells.put(cell.getColumnIndex(), new CellSnapshot(cell.getCellType(), value, cell.getCellStyle()));
        }
        return new RowSnapshot(row.getHeight(), cells);
    }

    private void restore(Row row, RowSnapshot snapshot) {
        List<Cell> existing = new ArrayList<>();
        row.forEach(existing::add);
        existing.forEach(row::removeCell);
        row.setHeight(snapshot.height());
        snapshot.cells().forEach((column, cellSnapshot) -> {
            Cell cell = row.createCell(column, cellSnapshot.type());
            if (cellSnapshot.style() != null) cell.setCellStyle(cellSnapshot.style());
            if (cellSnapshot.value() == null) return;
            switch (cellSnapshot.type()) {
                case STRING -> cell.setCellValue((String) cellSnapshot.value());
                case NUMERIC -> cell.setCellValue((Double) cellSnapshot.value());
                case BOOLEAN -> cell.setCellValue((Boolean) cellSnapshot.value());
                case FORMULA -> cell.setCellFormula((String) cellSnapshot.value());
                case ERROR -> cell.setCellErrorValue((Byte) cellSnapshot.value());
                default -> { }
            }
        });
    }

    private LocalDate periodDate(Cell cell, DataFormatter formatter) {
        if (cell == null) return null;
        if (cell.getCellType() == CellType.NUMERIC && DateUtil.isCellDateFormatted(cell)) {
            return cell.getLocalDateTimeCellValue().toLocalDate().withDayOfMonth(1);
        }
        return periodDate(formatter.formatCellValue(cell));
    }

    private LocalDate periodDate(String value) {
        if (value == null) return null;
        var matcher = java.util.regex.Pattern.compile("(?iu)^([\\p{L}]{3,4})\\.?-(\\d{2}|\\d{4})$")
                .matcher(value.trim());
        if (!matcher.matches()) return null;
        Integer month = Map.ofEntries(
                Map.entry("ene", 1), Map.entry("feb", 2), Map.entry("mar", 3), Map.entry("abr", 4),
                Map.entry("may", 5), Map.entry("jun", 6), Map.entry("jul", 7), Map.entry("ago", 8),
                Map.entry("sep", 9), Map.entry("sept", 9), Map.entry("oct", 10), Map.entry("nov", 11),
                Map.entry("dic", 12)
        ).get(matcher.group(1).toLowerCase(Locale.ROOT));
        if (month == null) return null;
        int year = Integer.parseInt(matcher.group(2));
        if (year < 100) year += 2000;
        return LocalDate.of(year, month, 1);
    }

    private boolean samePeriod(LocalDate left, LocalDate right) {
        return left != null && right != null
                && left.getYear() == right.getYear()
                && left.getMonth() == right.getMonth();
    }

    private String canonicalPeriodLabel(LocalDate date) {
        if (date == null) throw new IllegalArgumentException("El período mensual debe tener formato mmm-AA");
        String[] months = {"", "ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sep", "oct", "nov", "dic"};
        return months[date.getMonthValue()] + "-" + String.format("%02d", date.getYear() % 100);
    }

    private record CellSnapshot(CellType type, Object value, CellStyle style) { }
    private record RowSnapshot(short height, Map<Integer, CellSnapshot> cells) { }
    private record PeriodRowSnapshot(int rowIndex, LocalDate period, RowSnapshot row) { }
}
