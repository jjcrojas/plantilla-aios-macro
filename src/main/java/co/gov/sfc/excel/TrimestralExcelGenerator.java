package co.gov.sfc.excel;

import co.gov.sfc.config.AiosProperties;
import org.apache.poi.ss.usermodel.*;
import org.springframework.stereotype.Component;

import java.io.InputStream;
import java.math.BigDecimal;
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
public class TrimestralExcelGenerator {

    private final AiosProperties properties;

    public TrimestralExcelGenerator(AiosProperties properties) {
        this.properties = properties;
    }

    public Path generar(LocalDate fechaCorte, TrimestralData data) {
        return generar(List.of(new PeriodoTrimestral(fechaCorte, data)));
    }

    public Path generar(List<PeriodoTrimestral> periodos) {
        return generarDesdePlantilla(periodos, "Boletin_AIOS TRIMESTRAL.xlsx", "Boletin_AIOS TRIMESTRAL.xlsx");
    }

    public Path generarSemestral(LocalDate fechaCorte, TrimestralData data) {
        return generarDesdePlantilla(List.of(new PeriodoTrimestral(fechaCorte, data)),
                "Boletin_AIOS SEMESTRAL.xlsx", "Boletin_AIOS SEMESTRAL.xlsx");
    }

    private Path generarDesdePlantilla(List<PeriodoTrimestral> periodos, String plantillaNombre, String salidaNombre) {
        if (periodos == null || periodos.isEmpty()) {
            throw new IllegalArgumentException("Debe suministrar al menos un período trimestral");
        }
        Path base = properties.salidasReferenciaDir().resolve(plantillaNombre);
        if (!Files.isRegularFile(base) && "Boletin_AIOS SEMESTRAL.xlsx".equals(plantillaNombre)) {
            base = properties.salidasReferenciaDir().resolve("Boletin_AIOS TRIMESTRAL.xlsx");
        }
        Path outDir = Path.of("target", "aios-output");

        try {
            Files.createDirectories(outDir);
            Path out = Files.createTempDirectory(outDir, "trimestral-").resolve(salidaNombre);
            try (InputStream in = Files.newInputStream(base); Workbook wb = WorkbookFactory.create(in)) {
                periodos.stream()
                        .sorted(Comparator.comparing(PeriodoTrimestral::fechaCorte))
                        .forEach(periodo -> escribirPeriodo(wb, periodo));

                try (var os = Files.newOutputStream(out)) {
                    wb.write(os);
                }
            }
            return out;
        } catch (Exception e) {
            throw new IllegalStateException("No fue posible generar boletín " + salidaNombre, e);
        }
    }

    private void escribirPeriodo(Workbook wb, PeriodoTrimestral periodo) {
        LocalDate fechaCorte = periodo.fechaCorte();
        TrimestralData data = periodo.data();
        int filaAf = findOrAppendRow(wb.getSheet("afiliados"), fechaCorte, data.etiquetaFecha());
        int filaAport = findOrAppendRow(wb.getSheet("aportantes"), fechaCorte, data.etiquetaFecha());
        int filaCol = findOrAppendRow(wb.getSheet("colombia"), fechaCorte, data.etiquetaFecha());
        int filaTrasp = findOrAppendRow(wb.getSheet("traspasos"), fechaCorte, data.etiquetaFecha());
        int filaGast = findOrAppendRow(wb.getSheet("gastos"), fechaCorte, data.etiquetaFecha());
        int filaProm = findOrAppendRow(wb.getSheet("promotores"), fechaCorte, data.etiquetaFecha());
        int filaRent = findOrAppendRow(wb.getSheet("rentabilidad"), fechaCorte, data.etiquetaFecha());
        int filaCom = findOrAppendRow(wb.getSheet("comisiones"), fechaCorte, data.etiquetaFecha());

        writeAfiliados(wb.getSheet("afiliados"), filaAf, data.afiliados());
        writeAportantes(wb.getSheet("aportantes"), filaAport, data.aportantes());
        writeColombia(wb.getSheet("colombia"), filaCol, data.colombiaUsd());
        writeTraspasos(wb.getSheet("traspasos"), filaTrasp, data.traspasos());
        writeGastos(wb.getSheet("gastos"), filaGast, data.gastosUsd());
        writePromotores(wb.getSheet("promotores"), filaProm);
        writeRentabilidad(wb.getSheet("rentabilidad"), filaRent, data.rentNominalPct(), data.rentRealPct());
        writeComisiones(wb.getSheet("comisiones"), filaCom, data.comisionesPct());
    }

    private void writeAfiliados(Sheet s, int r, Map<String, BigDecimal> a) {
        write(s, r, 2, v(a, "mod_colf")); write(s, r, 3, v(a, "con_colf")); write(s, r, 4, v(a, "mr_colf"));
        write(s, r, 5, v(a, "con_mod_colf")); write(s, r, 6, v(a, "con_mr_colf")); write(s, r, 7, v(a, "mod_mr_colf"));
        write(s, r, 13, v(a, "mod_porv")); write(s, r, 14, v(a, "con_porv")); write(s, r, 15, v(a, "mr_porv"));
        write(s, r, 16, v(a, "con_mod_porv")); write(s, r, 17, v(a, "con_mr_porv")); write(s, r, 18, v(a, "mod_mr_porv"));
        write(s, r, 19, v(a, "mod_prot")); write(s, r, 20, v(a, "con_prot")); write(s, r, 21, v(a, "mr_prot"));
        write(s, r, 22, v(a, "con_mod_prot")); write(s, r, 23, v(a, "con_mr_prot")); write(s, r, 24, v(a, "mod_mr_prot"));
        for (int c = 25; c <= 29; c++) write(s, r, c, BigDecimal.ZERO);
        write(s, r, 30, v(a, "mod_sk_total")); write(s, r, 31, v(a, "con_sk")); write(s, r, 32, v(a, "mr_sk"));
        write(s, r, 33, v(a, "con_mod_sk")); write(s, r, 34, v(a, "con_mr_sk")); write(s, r, 35, v(a, "mod_mr_sk"));
    }

    private void writeAportantes(Sheet s, int r, Map<String, BigDecimal> a) {
        write(s, r, 2, v(a, "colf")); write(s, r, 3, BigDecimal.ZERO); write(s, r, 4, v(a, "porv"));
        write(s, r, 5, v(a, "prot")); write(s, r, 6, BigDecimal.ZERO); write(s, r, 7, v(a, "sk"));
    }

    private void writeColombia(Sheet s, int r, Map<String, BigDecimal> c) {
        writeText(s, r, 10, text(s, r, 1));  // J = misma etiqueta fecha
        writeText(s, r, 19, text(s, r, 1));  // S = misma etiqueta fecha
        writeText(s, r, 28, text(s, r, 1));  // AB = misma etiqueta fecha

        write(s, r, 2, v(c, "mod_colf")); write(s, r, 3, BigDecimal.ZERO); write(s, r, 4, v(c, "mod_porv"));
        write(s, r, 5, v(c, "mod_prot")); write(s, r, 6, BigDecimal.ZERO); write(s, r, 7, v(c, "mod_sk").add(v(c, "mod_alt")));
        write(s, r, 11, v(c, "con_colf")); write(s, r, 12, BigDecimal.ZERO); write(s, r, 13, v(c, "con_porv"));
        write(s, r, 14, v(c, "con_prot")); write(s, r, 15, BigDecimal.ZERO); write(s, r, 16, v(c, "con_sk"));
        write(s, r, 20, v(c, "mr_colf")); write(s, r, 21, BigDecimal.ZERO); write(s, r, 22, v(c, "mr_porv"));
        write(s, r, 23, v(c, "mr_prot")); write(s, r, 24, BigDecimal.ZERO); write(s, r, 25, v(c, "mr_sk"));
        write(s, r, 29, v(c, "rp_colf")); write(s, r, 30, BigDecimal.ZERO); write(s, r, 31, v(c, "rp_porv"));
        write(s, r, 32, v(c, "rp_prot")); write(s, r, 33, BigDecimal.ZERO); write(s, r, 34, v(c, "rp_sk"));
    }

    private void writeTraspasos(Sheet s, int r, Map<String, BigDecimal> t) {
        write(s, r, 2, v(t, "colf")); write(s, r, 3, BigDecimal.ZERO); write(s, r, 4, v(t, "porv"));
        write(s, r, 5, v(t, "prot")); write(s, r, 6, BigDecimal.ZERO); write(s, r, 7, v(t, "sk"));
    }

    private void writeGastos(Sheet s, int r, Map<String, BigDecimal> g) {
        write(s, r, 2, v(g, "colf")); write(s, r, 3, BigDecimal.ZERO); write(s, r, 4, v(g, "porv"));
        write(s, r, 5, v(g, "prot")); write(s, r, 6, BigDecimal.ZERO); write(s, r, 7, v(g, "sk"));
    }

    private void writePromotores(Sheet s, int r) { for (int c = 2; c <= 7; c++) writeText(s, r, c, "n.d."); }

    private void writeRentabilidad(Sheet s, int r, Map<String, BigDecimal> nom, Map<String, BigDecimal> real) {
        write(s, r, 2, v(nom, "colf")); write(s, r, 3, BigDecimal.ZERO); write(s, r, 4, v(nom, "porv"));
        write(s, r, 5, v(nom, "prot")); write(s, r, 6, BigDecimal.ZERO); write(s, r, 7, v(nom, "oldm"));
        write(s, r, 10, v(real, "colf")); write(s, r, 11, BigDecimal.ZERO); write(s, r, 12, v(real, "porv"));
        write(s, r, 13, v(real, "prot")); write(s, r, 14, BigDecimal.ZERO); write(s, r, 15, v(real, "oldm"));
    }

    private void writeComisiones(Sheet s, int r, Map<String, BigDecimal> c) {
        write(s, r, 2, v(c, "col_obl")); write(s, r, 3, v(c, "col_seg")); write(s, r, 4, BigDecimal.ZERO); write(s, r, 5, BigDecimal.ZERO);
        write(s, r, 6, v(c, "por_obl")); write(s, r, 7, v(c, "por_seg")); write(s, r, 8, v(c, "pro_obl")); write(s, r, 9, v(c, "pro_seg"));
        write(s, r, 10, BigDecimal.ZERO); write(s, r, 11, BigDecimal.ZERO); write(s, r, 12, v(c, "ska_obl")); write(s, r, 13, v(c, "ska_seg"));
    }

    private BigDecimal v(Map<String, BigDecimal> m, String k) { return m == null ? BigDecimal.ZERO : m.getOrDefault(k, BigDecimal.ZERO); }

    int findOrAppendRow(Sheet sheet, LocalDate fechaCorte, String etiqueta) {
        if (sheet == null) throw new IllegalStateException("No existe una hoja requerida en Boletin_AIOS TRIMESTRAL.xlsx");
        sortAndNormalizePeriodRows(sheet);
        DataFormatter formatter = new DataFormatter();
        String etiquetaCanonica = canonicalPeriodLabel(fechaCorte);
        for (int r = 0; r <= sheet.getLastRowNum(); r++) {
            Row row = sheet.getRow(r); if (row == null) continue;
            Cell c = row.getCell(0); if (c == null) continue;
            LocalDate periodo = periodDate(c, formatter);
            if (samePeriod(periodo, fechaCorte)) {
                c.setCellValue(etiquetaCanonica);
                normalizeSeptemberLabels(row, formatter);
                return r + 1;
            }
        }

        int lastPeriodRow = -1;
        int insertionRow = -1;
        for (int rowIndex = 5; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
            Row candidate = sheet.getRow(rowIndex);
            if (candidate == null) continue;
            LocalDate periodo = periodDate(candidate.getCell(0), formatter);
            if (periodo == null) continue;
            lastPeriodRow = rowIndex;
            if (insertionRow < 0 && periodo.isAfter(fechaCorte)) insertionRow = rowIndex;
        }
        int r = insertionRow >= 0 ? insertionRow : Math.max(lastPeriodRow + 1, 6);
        if (r <= sheet.getLastRowNum()) {
            sheet.shiftRows(r, sheet.getLastRowNum(), 1, true, false);
        }
        Row row = sheet.getRow(r); if (row == null) row = sheet.createRow(r);
        copyPreviousRowFormat(sheet, r);
        Cell c = row.getCell(0); if (c == null) c = row.createCell(0);
        c.setCellValue(etiquetaCanonica);
        return r + 1;
    }

    private void sortAndNormalizePeriodRows(Sheet sheet) {
        DataFormatter formatter = new DataFormatter();
        List<PeriodRowSnapshot> periods = new ArrayList<>();
        for (int rowIndex = 5; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
            Row row = sheet.getRow(rowIndex);
            if (row == null) continue;
            LocalDate period = periodDate(row.getCell(0), formatter);
            if (period != null) periods.add(new PeriodRowSnapshot(rowIndex, period, snapshot(row)));
        }
        if (periods.isEmpty()) return;

        List<RowSnapshot> ordered = periods.stream()
                .sorted(Comparator.comparing(PeriodRowSnapshot::period))
                .map(PeriodRowSnapshot::row)
                .toList();
        for (int i = 0; i < periods.size(); i++) {
            Row target = sheet.getRow(periods.get(i).rowIndex());
            if (target == null) target = sheet.createRow(periods.get(i).rowIndex());
            restore(target, ordered.get(i));
            normalizeSeptemberLabels(target, formatter);
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

    private void normalizeSeptemberLabels(Row row, DataFormatter formatter) {
        for (Cell cell : row) {
            String value = formatter.formatCellValue(cell);
            if (value != null && value.trim().matches("(?iu)^sept\\.?-\\d{2,4}$")) {
                cell.setCellValue(value.trim().replaceFirst("(?iu)^sept\\.?-", "sep-"));
            }
        }
    }

    private LocalDate periodDate(Cell cell, DataFormatter formatter) {
        if (cell == null) return null;
        if (cell.getCellType() == CellType.NUMERIC && DateUtil.isCellDateFormatted(cell)) {
            LocalDate date = cell.getLocalDateTimeCellValue().toLocalDate();
            return date.withDayOfMonth(1);
        }
        String value = formatter.formatCellValue(cell);
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
        return left != null && left.getYear() == right.getYear() && left.getMonth() == right.getMonth();
    }

    private String canonicalPeriodLabel(LocalDate date) {
        String[] months = {"", "ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sep", "oct", "nov", "dic"};
        return months[date.getMonthValue()] + "-" + String.format("%02d", date.getYear() % 100);
    }

    private record CellSnapshot(CellType type, Object value, CellStyle style) { }
    private record RowSnapshot(short height, Map<Integer, CellSnapshot> cells) { }
    private record PeriodRowSnapshot(int rowIndex, LocalDate period, RowSnapshot row) { }
    public record PeriodoTrimestral(LocalDate fechaCorte, TrimestralData data) { }

    private void copyPreviousRowFormat(Sheet sheet, int targetRowIndex) {
        if (targetRowIndex <= 0) return;
        Row source = sheet.getRow(targetRowIndex - 1);
        Row target = sheet.getRow(targetRowIndex);
        if (source == null || target == null) return;

        target.setHeight(source.getHeight());
        int lastCell = Math.max(source.getLastCellNum(), 1);
        for (int col = 0; col < lastCell; col++) {
            Cell sourceCell = source.getCell(col);
            if (sourceCell == null) continue;
            Cell targetCell = target.getCell(col);
            if (targetCell == null) targetCell = target.createCell(col);
            targetCell.setCellStyle(sourceCell.getCellStyle());
        }
    }

    private void write(Sheet sheet, int row1, int col1, BigDecimal value) {
        Row row = sheet.getRow(row1 - 1); if (row == null) row = sheet.createRow(row1 - 1);
        Cell cell = row.getCell(col1 - 1); if (cell == null) cell = row.createCell(col1 - 1);
        cell.setCellValue(value == null ? 0d : value.doubleValue());
    }

    private void writeText(Sheet sheet, int row1, int col1, String value) {
        Row row = sheet.getRow(row1 - 1); if (row == null) row = sheet.createRow(row1 - 1);
        Cell cell = row.getCell(col1 - 1); if (cell == null) cell = row.createCell(col1 - 1);
        cell.setCellValue(value);
    }

    private String text(Sheet sheet, int row1, int col1) {
        Row row = sheet.getRow(row1 - 1);
        if (row == null) return "";
        Cell cell = row.getCell(col1 - 1);
        if (cell == null) return "";
        return new DataFormatter().formatCellValue(cell);
    }
}
