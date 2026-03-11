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

@Component
public class MensualExcelGenerator {

    private final AiosProperties properties;
    private final CeldaLogger celdaLogger;

    public MensualExcelGenerator(AiosProperties properties, CeldaLogger celdaLogger) {
        this.properties = properties;
        this.celdaLogger = celdaLogger;
    }

    public Path generar(MensualData data) {
        Path baseMensual = properties.salidasReferenciaDir().resolve("Boletin_AIOS MENSUAL.xlsx");
        Path outDir = Path.of("target", "aios-output");
        Path out = outDir.resolve("Boletin_AIOS MENSUAL.xlsx");

        try {
            Files.createDirectories(outDir);
            try (InputStream in = Files.newInputStream(baseMensual); Workbook wb = WorkbookFactory.create(in)) {
                Sheet sheet = wb.getSheet("HOJA1");
                int row = findDateRow(sheet, data.textoFecha());
                write(sheet, row, 2, data.afiliados(), "Serie_Formato_ 491 AFILIADOS AFP.xlsm", "informe de prensa", "C11+D11");
                write(sheet, row, 3, data.aportantes(), "Serie_Formato_ 491 AFILIADOS AFP.xlsm", "multifondos", "E25");
                write(sheet, row, 4, data.traspasosSistema(), "Serie_Formato_493 MOVIMIENTO AFILIADOS.xlsx", "Traslados Entre AFP", "BQ11");
                write(sheet, row, 5, divide(data.vrFondo(), trm(data)), "SISTEMA TOTAL *.xls", "restot", "SISTEMA/MAX");
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
                write(sheet, row, 17, data.consFdosAdmon(), "Serie_Formato_ 491 AFILIADOS AFP.xlsm", "multifondos", "J8+J9/J12");
                write(sheet, row, 18, data.porcVrFondo(), "SISTEMA TOTAL *.xls", "restot", "(PROTECCION+PORVENIR)/SISTEMA/10");
                write(sheet, row, 19, trm(data), "series PIB_PEA_TRM_DG.xlsm", "Hoja1", "TRM<=fecha");

                try (OutputStream os = Files.newOutputStream(out)) {
                    wb.write(os);
                }
            }
            return out;
        } catch (Exception e) {
            throw new IllegalStateException("No fue posible generar boletín mensual", e);
        }
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

    private void write(Sheet sheet, int row1Based, int col1Based, BigDecimal value, String fuenteArchivo, String fuenteHoja, String fuenteCelda) {
        Row row = sheet.getRow(row1Based - 1);
        if (row == null) row = sheet.createRow(row1Based - 1);
        Cell cell = row.getCell(col1Based - 1);
        if (cell == null) cell = row.createCell(col1Based - 1, CellType.NUMERIC);
        cell.setCellValue(value.doubleValue());
        String celda = CellReference.convertNumToColString(col1Based - 1) + row1Based;
        celdaLogger.log(sheet.getSheetName(), celda, value, fuenteArchivo, fuenteHoja, fuenteCelda);
    }

    private int findDateRow(Sheet sheet, String textoFecha) {
        DataFormatter formatter = new DataFormatter();
        for (Row row : sheet) {
            String value = formatter.formatCellValue(row.getCell(0));
            if (value != null && value.trim().equalsIgnoreCase(textoFecha.trim())) {
                return row.getRowNum() + 1;
            }
        }
        throw new IllegalArgumentException("No se encontró la fecha " + textoFecha + " en HOJA1 columna A.");
    }
}
