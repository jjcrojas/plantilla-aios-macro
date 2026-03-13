package co.gov.sfc.excel;

import co.gov.sfc.config.AiosProperties;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.springframework.stereotype.Component;

import java.io.InputStream;
import java.math.BigDecimal;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.LocalDate;

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

                // afiliados: sin desglose por AFP en la versión Java (se preserva estructura de macro con ceros).
                Sheet afiliados = wb.getSheet("afiliados");
                for (int c = 2; c <= 35; c++) write(afiliados, filaAf, c, BigDecimal.ZERO);

                // aportantes (491): Colfondos, Porvenir, Protección, Skandia.
                Sheet aportantes = wb.getSheet("aportantes");
                write(aportantes, filaAport, 2, data.cotColfondos());
                write(aportantes, filaAport, 3, BigDecimal.ZERO);
                write(aportantes, filaAport, 4, data.cotPorvenir());
                write(aportantes, filaAport, 5, data.cotProteccion());
                write(aportantes, filaAport, 6, BigDecimal.ZERO);
                write(aportantes, filaAport, 7, data.cotSkandia());

                // colombia: por ahora se ubica el saldo total en Colfondos moderado y el resto en cero.
                Sheet colombia = wb.getSheet("colombia");
                write(colombia, filaCol, 2, data.vrFondoUsd());
                for (int c : new int[]{3, 4, 5, 6, 7, 9, 10, 11, 12, 13, 14, 16, 17, 18, 19, 20, 21, 23, 24, 25, 26, 27, 28}) {
                    write(colombia, filaCol, c, BigDecimal.ZERO);
                }

                // traspasos: por ahora solo total del sistema en la primera columna de AFP.
                Sheet traspasos = wb.getSheet("traspasos");
                write(traspasos, filaTrasp, 2, data.traspasosSistema());
                write(traspasos, filaTrasp, 3, BigDecimal.ZERO);
                write(traspasos, filaTrasp, 4, BigDecimal.ZERO);
                write(traspasos, filaTrasp, 5, BigDecimal.ZERO);
                write(traspasos, filaTrasp, 6, BigDecimal.ZERO);
                write(traspasos, filaTrasp, 7, BigDecimal.ZERO);

                Sheet gastos = wb.getSheet("gastos");
                for (int c = 2; c <= 7; c++) write(gastos, filaGast, c, BigDecimal.ZERO);

                Sheet promotores = wb.getSheet("promotores");
                for (int c = 2; c <= 7; c++) writeText(promotores, filaProm, c, "n.d.");

                Sheet rentabilidad = wb.getSheet("rentabilidad");
                write(rentabilidad, filaRent, 2, data.rentNominal12m());
                write(rentabilidad, filaRent, 3, BigDecimal.ZERO);
                write(rentabilidad, filaRent, 4, BigDecimal.ZERO);
                write(rentabilidad, filaRent, 5, BigDecimal.ZERO);
                write(rentabilidad, filaRent, 6, BigDecimal.ZERO);
                write(rentabilidad, filaRent, 7, BigDecimal.ZERO);
                write(rentabilidad, filaRent, 10, data.rentReal12m());
                write(rentabilidad, filaRent, 11, BigDecimal.ZERO);
                write(rentabilidad, filaRent, 12, BigDecimal.ZERO);
                write(rentabilidad, filaRent, 13, BigDecimal.ZERO);
                write(rentabilidad, filaRent, 14, BigDecimal.ZERO);
                write(rentabilidad, filaRent, 15, BigDecimal.ZERO);

                Sheet comisiones = wb.getSheet("comisiones");
                for (int c = 2; c <= 13; c++) write(comisiones, filaCom, c, BigDecimal.ZERO);

                try (var os = Files.newOutputStream(out)) {
                    wb.write(os);
                }
            }
            return out;
        } catch (Exception e) {
            throw new IllegalStateException("No fue posible generar boletín trimestral", e);
        }
    }

    private int findOrAppendRow(Sheet sheet, LocalDate fechaCorte, String etiqueta) {
        if (sheet == null) {
            throw new IllegalStateException("No existe una hoja requerida en Boletin_AIOS TRIMESTRAL.xlsx");
        }
        DataFormatter formatter = new DataFormatter();
        String etiquetaNorm = etiqueta == null ? "" : etiqueta.trim().toLowerCase();

        for (int r = 0; r <= sheet.getLastRowNum(); r++) {
            Row row = sheet.getRow(r);
            if (row == null) continue;
            Cell c = row.getCell(0);
            if (c == null) continue;

            String texto = formatter.formatCellValue(c);
            if (texto != null && texto.trim().toLowerCase().equals(etiquetaNorm)) {
                return r + 1;
            }

            if (c.getCellType() == CellType.NUMERIC && DateUtil.isCellDateFormatted(c)) {
                LocalDate d = c.getLocalDateTimeCellValue().toLocalDate();
                if (d.getYear() == fechaCorte.getYear() && d.getMonth() == fechaCorte.getMonth()) {
                    return r + 1;
                }
            }
        }

        int r = Math.max(sheet.getLastRowNum() + 1, 6);
        Row row = sheet.getRow(r);
        if (row == null) row = sheet.createRow(r);
        Cell c = row.getCell(0);
        if (c == null) c = row.createCell(0);
        c.setCellValue(etiqueta);
        return r + 1;
    }

    private void write(Sheet sheet, int row1, int col1, BigDecimal value) {
        Row row = sheet.getRow(row1 - 1);
        if (row == null) row = sheet.createRow(row1 - 1);
        Cell cell = row.getCell(col1 - 1);
        if (cell == null) cell = row.createCell(col1 - 1);
        cell.setCellValue(value == null ? 0d : value.doubleValue());
    }

    private void writeText(Sheet sheet, int row1, int col1, String value) {
        Row row = sheet.getRow(row1 - 1);
        if (row == null) row = sheet.createRow(row1 - 1);
        Cell cell = row.getCell(col1 - 1);
        if (cell == null) cell = row.createCell(col1 - 1);
        cell.setCellValue(value);
    }
}
