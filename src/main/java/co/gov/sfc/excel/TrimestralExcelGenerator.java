package co.gov.sfc.excel;

import co.gov.sfc.config.AiosProperties;
import org.apache.poi.ss.usermodel.*;
import org.springframework.stereotype.Component;

import java.io.InputStream;
import java.math.BigDecimal;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.LocalDate;
import java.util.Map;

@Component
public class TrimestralExcelGenerator {

    private final AiosProperties properties;

    public TrimestralExcelGenerator(AiosProperties properties) {
        this.properties = properties;
    }

    public Path generar(LocalDate fechaCorte, TrimestralData data) {
        Path base = properties.salidasReferenciaDir().resolve("Boletin_AIOS TRIMESTRAL.xlsx");
        Path outDir = Path.of("target", "aios-output");
        Path out = outDir.resolve("Boletin_AIOS TRIMESTRAL.xlsx");

        try {
            Files.createDirectories(outDir);
            try (InputStream in = Files.newInputStream(base); Workbook wb = WorkbookFactory.create(in)) {
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

                try (var os = Files.newOutputStream(out)) {
                    wb.write(os);
                }
            }
            return out;
        } catch (Exception e) {
            throw new IllegalStateException("No fue posible generar boletín trimestral", e);
        }
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

    private int findOrAppendRow(Sheet sheet, LocalDate fechaCorte, String etiqueta) {
        if (sheet == null) throw new IllegalStateException("No existe una hoja requerida en Boletin_AIOS TRIMESTRAL.xlsx");
        DataFormatter formatter = new DataFormatter();
        String etiquetaNorm = etiqueta == null ? "" : etiqueta.trim().toLowerCase();
        for (int r = 0; r <= sheet.getLastRowNum(); r++) {
            Row row = sheet.getRow(r); if (row == null) continue;
            Cell c = row.getCell(0); if (c == null) continue;
            String texto = formatter.formatCellValue(c);
            if (texto != null && texto.trim().toLowerCase().equals(etiquetaNorm)) return r + 1;
            if (c.getCellType() == CellType.NUMERIC && DateUtil.isCellDateFormatted(c)) {
                LocalDate d = c.getLocalDateTimeCellValue().toLocalDate();
                if (d.getYear() == fechaCorte.getYear() && d.getMonth() == fechaCorte.getMonth()) return r + 1;
            }
        }
        int r = Math.max(sheet.getLastRowNum() + 1, 6);
        Row row = sheet.getRow(r); if (row == null) row = sheet.createRow(r);
        Cell c = row.getCell(0); if (c == null) c = row.createCell(0);
        c.setCellValue(etiqueta);
        return r + 1;
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
