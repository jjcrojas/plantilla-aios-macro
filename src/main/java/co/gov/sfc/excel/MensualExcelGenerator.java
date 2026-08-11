package co.gov.sfc.excel;

import co.gov.sfc.config.AiosProperties;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellReference;
import org.springframework.stereotype.Component;

import java.io.InputStream;
import java.io.OutputStream;
import java.math.BigDecimal;
import java.math.RoundingMode;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;

@Component
public class MensualExcelGenerator {

    private final AiosProperties properties;
    private final CeldaLogger celdaLogger;

    public MensualExcelGenerator(AiosProperties properties, CeldaLogger celdaLogger) {
        this.properties = properties;
        this.celdaLogger = celdaLogger;
    }

    public Path generar(MensualData data) {
        return generar(List.of(data));
    }

    public Path generar(List<MensualData> datos) {
        if (datos == null || datos.isEmpty()) {
            throw new IllegalArgumentException("Debe suministrar al menos un período mensual");
        }
        Path baseMensual = properties.salidasReferenciaDir().resolve("Boletin_AIOS MENSUAL.xlsx");
        Path outDir = Path.of("target", "aios-output");
        Path out = outDir.resolve("Boletin_AIOS MENSUAL.xlsx");

        try {
            Files.createDirectories(outDir);
            try (InputStream in = Files.newInputStream(baseMensual); Workbook wb = WorkbookFactory.create(in)) {
                Sheet sheet = wb.getSheet("HOJA1");
                for (MensualData data : datos) {
                    escribirPeriodo(sheet, data);
                }

                try (OutputStream os = Files.newOutputStream(out)) {
                    wb.write(os);
                }
            }
            return out;
        } catch (Exception e) {
            throw new IllegalStateException("No fue posible generar boletín mensual", e);
        }
    }

    private void escribirPeriodo(Sheet sheet, MensualData data) {
        int row = findOrCreateDateRow(sheet, data.textoFecha());
        aplicarFormatoFilaMensual(sheet, row);
        write(sheet, row, 2, data.afiliados(), "Query Teradata PROD_DWH_CONSULTA.FORMATO491", "FORMATO491", "SUM(TOTAL_AFILIADOS_TOTAL), RENGLON=999");
                write(sheet, row, 3, data.aportantes(), "Query Teradata PROD_DWH_CONSULTA.FORMATO491", "FORMATO491", "SUM(TOTAL_AFILIADOS_COTIZANTES), RENGLON=999");
                write(sheet, row, 4, data.traspasosSistema(), "Query Teradata PROD_DWH_CONSULTA.S9_FORMATO_493", "S9_FORMATO_493", "Traspasos sistema 12 meses por FECHA_CORTE/UNIDAD_CAPTURA/RENGLON");
                write(sheet, row, 5, divide(data.vrFondo(), trm(data)), "Query Teradata PROD_DWH_CONSULTA.ESTFIN_INDIV_PA", "PUC 100000", "SUM(Saldo_Sincierre_Total_Moneda_0)/1000000/TRM");
                write(sheet, row, 6, divide(data.total1(), trm(data)), "LIMITES del nuevo.xlsm", "AIOS", "AB4");
                write(sheet, row, 7, pct(data.dudaG()), "LIMITES del nuevo.xlsm", "AIOS", "C4");
                write(sheet, row, 8, pct(data.dudaEf()), "LIMITES del nuevo.xlsm", "AIOS", "E4");
                write(sheet, row, 9, pct(data.dudaNf()), "LIMITES del nuevo.xlsm", "AIOS", "G4");
                write(sheet, row, 10, pct(data.dudaAc()), "LIMITES del nuevo.xlsm", "AIOS", "I4");
                write(sheet, row, 11, pct(data.dudaF()), "LIMITES del nuevo.xlsm", "AIOS", "K4");
                write(sheet, row, 12, pct(data.h17()), "LIMITES del nuevo.xlsm", "AIOS", "O4:Y4");
                write(sheet, row, 13, pct(data.otros()), "LIMITES del nuevo.xlsm", "AIOS", "AA4");
                write(sheet, row, 14, pct(data.tmpNominal1()), "Rent_Vr_Uni_Moderado.xlsm", "(primera)", "D11");
                write(sheet, row, 15, pct(data.tmpReal1()), "Rent_Vr_Uni_Moderado.xlsm", "(primera)", "D10");
                write(sheet, row, 16, BigDecimal.valueOf(4), "constante", "", "");
                write(sheet, row, 17, data.consFdosAdmon(), "Query Teradata PROD_DWH_CONSULTA.FORMATO491", "FORMATO491", "Top 2 AFP por TOTAL_AFILIADOS_TOTAL / total sistema");
                write(sheet, row, 18, data.porcVrFondo(), "Query Teradata PROD_DWH_CONSULTA.ESTFIN_INDIV_PA", "PUC 100000", "(Proteccion+Porvenir)/total sistema * 100");
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
        DataFormatter formatter = new DataFormatter();
        for (Row row : sheet) {
            String value = formatter.formatCellValue(row.getCell(0));
            if (value != null && value.trim().equalsIgnoreCase(textoFecha.trim())) {
                return row.getRowNum() + 1;
            }
        }

        int lastPeriodRow = 0;
        for (int rowIndex = 1; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
            Row candidate = sheet.getRow(rowIndex);
            if (candidate == null) continue;
            if (isPeriodCell(candidate.getCell(0), formatter)) lastPeriodRow = rowIndex;
        }
        int rowIndex = Math.max(lastPeriodRow + 1, 1);
        if (rowIndex <= sheet.getLastRowNum()) {
            sheet.shiftRows(rowIndex, sheet.getLastRowNum(), 1, true, false);
        }
        Row row = sheet.getRow(rowIndex);
        if (row == null) row = sheet.createRow(rowIndex);
        Cell dateCell = row.getCell(0);
        if (dateCell == null) dateCell = row.createCell(0, CellType.STRING);
        dateCell.setCellValue(textoFecha);
        return rowIndex + 1;
    }

    private boolean isPeriodCell(Cell cell, DataFormatter formatter) {
        if (cell == null) return false;
        if (cell.getCellType() == CellType.NUMERIC && DateUtil.isCellDateFormatted(cell)) return true;
        String value = formatter.formatCellValue(cell);
        return value != null && value.trim().matches("(?iu)^[\\p{L}]{3}\\.?-\\d{2,4}$");
    }
}
