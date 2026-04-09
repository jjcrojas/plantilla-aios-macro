package co.gov.sfc.excel;

import co.gov.sfc.config.AiosProperties;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.springframework.stereotype.Component;

import java.io.InputStream;
import java.math.BigDecimal;
import java.math.RoundingMode;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.LocalDate;
import java.time.YearMonth;

@Component
public class SemestralExcelGenerator {

    private static final YearMonth BASE_PERIODO = YearMonth.of(2017, 12);
    private static final int BASE_COLUMNA = 3;

    private final AiosProperties properties;

    public SemestralExcelGenerator(AiosProperties properties) {
        this.properties = properties;
    }

    public Path generar(LocalDate fechaCorte, MensualData mensual, TrimestralData trimestral) {
        Path base = resolveTemplate();
        Path outDir = Path.of("target", "aios-output");
        Path out = outDir.resolve("semestral.xlsx");

        try {
            Files.createDirectories(outDir);
            try (InputStream in = Files.newInputStream(base); Workbook wb = WorkbookFactory.create(in)) {
                Sheet hoja = resolveSheet(wb);
                int col = columnaSemestral(fechaCorte);

                // Bloque A - principales (según EscribirSemestral_Integral)
                write(hoja, 3, col, mensual.afiliados());
                write(hoja, 11, col, mensual.aportantes());
                write(hoja, 26, col, mensual.traspasosSistema());
                write(hoja, 28, col, divide(mensual.vrFondo(), trm(mensual)));

                // Bloque B - límites
                write(hoja, 30, col, divide(mensual.total1(), trm(mensual)));
                write(hoja, 31, col, mensual.dudaG());
                write(hoja, 32, col, mensual.dudaEf());
                write(hoja, 33, col, mensual.dudaNf());
                write(hoja, 34, col, mensual.dudaAc());
                write(hoja, 35, col, mensual.dudaF());
                write(hoja, 43, col, mensual.otros());
                write(hoja, 44, col, mensual.h17());

                // Bloque C/D - uso de datos disponibles del flujo Java
                BigDecimal gastosUsdTotal = trimestral.gastosUsd().values().stream().reduce(BigDecimal.ZERO, BigDecimal::add);
                write(hoja, 52, col, gastosUsdTotal);
                write(hoja, 71, col, promedioComisionObligatoria(trimestral).multiply(BigDecimal.valueOf(100)));
                write(hoja, 82, col, mensual.tmpNominal1().multiply(BigDecimal.valueOf(100)));
                write(hoja, 83, col, mensual.tmpReal1().multiply(BigDecimal.valueOf(100)));

                try (var os = Files.newOutputStream(out)) {
                    wb.write(os);
                }
            }
            return out;
        } catch (Exception e) {
            throw new IllegalStateException("No fue posible generar archivo semestral", e);
        }
    }

    private BigDecimal promedioComisionObligatoria(TrimestralData trimestral) {
        BigDecimal col = trimestral.comisionesPct().getOrDefault("col_obl", BigDecimal.ZERO);
        BigDecimal por = trimestral.comisionesPct().getOrDefault("por_obl", BigDecimal.ZERO);
        BigDecimal pro = trimestral.comisionesPct().getOrDefault("pro_obl", BigDecimal.ZERO);
        BigDecimal ska = trimestral.comisionesPct().getOrDefault("ska_obl", BigDecimal.ZERO);
        return col.add(por).add(pro).add(ska).divide(BigDecimal.valueOf(4), 8, RoundingMode.HALF_UP);
    }

    private Path resolveTemplate() {
        String[] nombres = {"semestral.xlsx", "Semestral_Colombia.xlsx", "Boletin_AIOS SEMESTRAL.xlsx"};
        for (String nombre : nombres) {
            Path candidate = properties.salidasReferenciaDir().resolve(nombre);
            if (Files.isRegularFile(candidate)) {
                return candidate;
            }
        }
        throw new IllegalStateException("No se encontró plantilla semestral en salidas_referencia");
    }

    private Sheet resolveSheet(Workbook wb) {
        Sheet s = wb.getSheet("Hoja1");
        if (s != null) return s;
        s = wb.getSheet("Hoja");
        if (s != null) return s;
        return wb.getSheetAt(0);
    }

    private int columnaSemestral(LocalDate fechaCorte) {
        int month = fechaCorte.getMonthValue();
        if (month != 6 && month != 12) {
            throw new IllegalArgumentException("La generación semestral solo aplica para junio o diciembre");
        }
        YearMonth period = YearMonth.of(fechaCorte.getYear(), month);
        int baseIndex = BASE_PERIODO.getYear() * 2 + (BASE_PERIODO.getMonthValue() == 12 ? 1 : 0);
        int targetIndex = period.getYear() * 2 + (period.getMonthValue() == 12 ? 1 : 0);
        int offset = targetIndex - baseIndex;
        if (offset < 0) {
            throw new IllegalArgumentException("La plantilla semestral no soporta periodos anteriores a diciembre de 2017");
        }
        return BASE_COLUMNA + offset;
    }

    private BigDecimal trm(MensualData data) {
        return data.trm().signum() == 0 ? BigDecimal.ONE : data.trm();
    }

    private BigDecimal divide(BigDecimal a, BigDecimal b) {
        if (b.signum() == 0) return BigDecimal.ZERO;
        return a.divide(b, 8, RoundingMode.HALF_UP);
    }

    private void write(Sheet sheet, int row1Based, int col1Based, BigDecimal value) {
        Row row = sheet.getRow(row1Based - 1);
        if (row == null) row = sheet.createRow(row1Based - 1);
        Cell cell = row.getCell(col1Based - 1);
        if (cell == null) cell = row.createCell(col1Based - 1);
        cell.setCellValue(value == null ? 0d : value.doubleValue());
    }
}
