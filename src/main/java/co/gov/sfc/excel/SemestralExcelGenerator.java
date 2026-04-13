package co.gov.sfc.excel;

import co.gov.sfc.config.AiosProperties;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Component;

import java.io.InputStream;
import java.math.BigDecimal;
import java.math.RoundingMode;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.LocalDate;
import java.util.Locale;

@Component
public class SemestralExcelGenerator {

    private static final Logger log = LoggerFactory.getLogger(SemestralExcelGenerator.class);

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
                int col = columnaSemestral(hoja, fechaCorte);

                // Bloque A - principales (según EscribirSemestral_Integral)
                write(hoja, 3, col, mensual.afiliados());
                write(hoja, 4, col, pct(safeDivide(mensual.afiliadosMenor30(), mensual.afiliados())));
                write(hoja, 5, col, pct(safeDivide(mensual.afiliados30a44(), mensual.afiliados())));
                write(hoja, 6, col, pct(safeDivide(mensual.afiliados45a59(), mensual.afiliados())));
                write(hoja, 7, col, pct(safeDivide(mensual.afiliadosMayor60(), mensual.afiliados())));
                write(hoja, 8, col, BigDecimal.valueOf(100));
                write(hoja, 9, col, divide(mensual.afiliados(), BigDecimal.valueOf(1000)));
                write(hoja, 10, col, pct(safeDivide(mensual.mujeres(), mensual.afiliados())));
                write(hoja, 11, col, mensual.aportantes());
                write(hoja, 12, col, pct(safeDivide(mensual.afiliados(), mensual.pea())));
                write(hoja, 13, col, pct(safeDivide(mensual.aportantes(), mensual.pea())));
                write(hoja, 14, col, pct(safeDivide(mensual.aportantes(), mensual.afiliados())));
                write(hoja, 15, col, mensual.smColombiaUsd());
                write(hoja, 16, col, mensual.totalPen());
                write(hoja, 17, col, safeDivide(mensual.totalInv(), mensual.totalPen()));
                write(hoja, 18, col, safeDivide(mensual.totalVej(), mensual.totalPen()));
                write(hoja, 19, col, safeDivide(mensual.totalSob(), mensual.totalPen()));
                log.info("Semestral: fila16(total_pen)={}, fila17(inv%)={}, fila18(vej%)={}, fila19(sob%)={} para fecha={} col={}.",
                        mensual.totalPen(),
                        safeDivide(mensual.totalInv(), mensual.totalPen()),
                        safeDivide(mensual.totalVej(), mensual.totalPen()),
                        safeDivide(mensual.totalSob(), mensual.totalPen()),
                        fechaCorte, col);
                write(hoja, 26, col, mensual.traspasosSistema());
                write(hoja, 27, col, safeDivide(mensual.traspasosSistema(), mensual.afiliados()));
                setNumberFormat(hoja, 27, col, "#,##0.00%");
                BigDecimal fondoCop = mensual.fondoSistemaJ14().multiply(BigDecimal.valueOf(1000));
                BigDecimal fondoUsdMM = safeDivide(safeDivide(fondoCop, trm(mensual)), BigDecimal.valueOf(1_000_000));
                write(hoja, 28, col, fondoUsdMM);
                BigDecimal ratioFondosPib = safeDivide(fondoCop, mensual.pibSemestral());
                write(hoja, 29, col, ratioFondosPib);
                setNumberFormat(hoja, 29, col, "#,##0.00%");
                log.info("Semestral traza fila29: fondoCop={} pibSemestral={} ratioFondosPib={} fecha={} col={}",
                        fondoCop, mensual.pibSemestral(), ratioFondosPib, fechaCorte, col);
                if (mensual.pibSemestral() == null || mensual.pibSemestral().signum() == 0) {
                    log.warn("Semestral fila29 en 0 por PIB nulo/cero; en la plantilla puede mostrarse '-' por formato contable. fecha={} col={}",
                            fechaCorte, col);
                }

                // Bloque B - límites
                write(hoja, 30, col, divide(mensual.total1(), trm(mensual)));
                write(hoja, 31, col, mensual.dudaG());
                write(hoja, 32, col, mensual.dudaEf());
                write(hoja, 33, col, mensual.dudaNf());
                write(hoja, 34, col, mensual.dudaAc());
                write(hoja, 35, col, mensual.dudaF());
                write(hoja, 36, col, BigDecimal.ZERO);
                write(hoja, 37, col, mensual.dudaGe());
                write(hoja, 38, col, mensual.dudaEfe());
                write(hoja, 39, col, mensual.dudaNfe());
                write(hoja, 40, col, mensual.dudaAce());
                write(hoja, 41, col, mensual.dudaFe());
                write(hoja, 42, col, BigDecimal.valueOf(2));
                write(hoja, 43, col, mensual.otros());
                write(hoja, 44, col, mensual.h17());
                BigDecimal deudaGobUsd = safeDivide(mensual.deudaGobB4().multiply(BigDecimal.valueOf(1_000_000)), trm(mensual));
                write(hoja, 45, col, safeDivide(safeDivide(fondoCop, trm(mensual)), deudaGobUsd));
                write(hoja, 46, col, BigDecimal.valueOf(4));
                write(hoja, 47, col, mensual.porcVrFondo());
                BigDecimal activos = mensual.activosCuentas() == null ? BigDecimal.ZERO : mensual.activosCuentas();
                BigDecimal pasivos = mensual.pasivosCuentas() == null ? BigDecimal.ZERO : mensual.pasivosCuentas();
                BigDecimal activosUsd = safeDivide(activos, trm(mensual));
                BigDecimal pasivosUsd = safeDivide(pasivos, trm(mensual));
                BigDecimal patrimonioUsd = safeDivide(activos.subtract(pasivos), trm(mensual));
                write(hoja, 48, col, activosUsd);
                write(hoja, 49, col, pasivosUsd);
                write(hoja, 50, col, patrimonioUsd);
                setNumberFormat(hoja, 48, col, "#,##0.00");
                setNumberFormat(hoja, 49, col, "#,##0.00");
                setNumberFormat(hoja, 50, col, "#,##0.00");
                log.info("Semestral traza filas48-50: activosCuentas(MM COP)={} pasivosCuentas(MM COP)={} trm={} -> activosUsd(MM USD)={} pasivosUsd(MM USD)={} patrimonioUsd(MM USD)={}",
                        activos, pasivos, trm(mensual), activosUsd, pasivosUsd, patrimonioUsd);

                // Bloque C/D - uso de datos disponibles del flujo Java
                BigDecimal gastosUsdTotal = trimestral.gastosUsd().values().stream().reduce(BigDecimal.ZERO, BigDecimal::add);
                write(hoja, 52, col, gastosUsdTotal);
                write(hoja, 71, col, promedioComisionObligatoria(trimestral).multiply(BigDecimal.valueOf(100)));
                write(hoja, 82, col, mensual.tmpNominal1().multiply(BigDecimal.valueOf(100)));
                write(hoja, 83, col, mensual.tmpReal1().multiply(BigDecimal.valueOf(100)));

                try (var os = Files.newOutputStream(out)) {
                    wb.write(os);
                }
                log.info("Semestral generado correctamente: archivo={} fecha={} columnaDestino={}", out.toAbsolutePath(), fechaCorte, col);
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

    private int columnaSemestral(Sheet hoja, LocalDate fechaCorte) {
        int month = fechaCorte.getMonthValue();
        if (month != 6 && month != 12) {
            throw new IllegalArgumentException("La generación semestral solo aplica para junio o diciembre");
        }

        String mesObjetivo = month == 6 ? "junio" : "diciembre";
        String anioObjetivo = String.valueOf(fechaCorte.getYear());
        DataFormatter fmt = new DataFormatter(Locale.forLanguageTag("es-CO"));

        Row rowMes = hoja.getRow(0);
        Row rowAnio = hoja.getRow(1);
        if (rowMes == null || rowAnio == null) {
            throw new IllegalStateException("La plantilla semestral no contiene encabezados de periodo en filas 1 y 2");
        }

        int last = Math.max(rowMes.getLastCellNum(), rowAnio.getLastCellNum());
        for (int c = 2; c < Math.max(last, 3); c++) {
            String mes = normalize(fmt.formatCellValue(rowMes.getCell(c)));
            String anio = normalize(fmt.formatCellValue(rowAnio.getCell(c))).replace(".0", "");
            if (mes.equals(mesObjetivo) && anio.equals(anioObjetivo)) {
                return c + 1;
            }
        }

        throw new IllegalArgumentException("No se encontró la columna para " + mesObjetivo + " " + anioObjetivo + " en la plantilla semestral");
    }

    private String normalize(String value) {
        return value == null ? "" : value.trim().toLowerCase(Locale.ROOT);
    }

    private BigDecimal trm(MensualData data) {
        return data.trm().signum() == 0 ? BigDecimal.ONE : data.trm();
    }

    private BigDecimal divide(BigDecimal a, BigDecimal b) {
        if (b.signum() == 0) return BigDecimal.ZERO;
        return a.divide(b, 8, RoundingMode.HALF_UP);
    }

    private BigDecimal safeDivide(BigDecimal numerator, BigDecimal denominator) {
        if (denominator == null || denominator.signum() == 0) {
            return BigDecimal.ZERO;
        }
        return (numerator == null ? BigDecimal.ZERO : numerator).divide(denominator, 8, RoundingMode.HALF_UP);
    }

    private BigDecimal pct(BigDecimal value) {
        return (value == null ? BigDecimal.ZERO : value).multiply(BigDecimal.valueOf(100));
    }

    private void write(Sheet sheet, int row1Based, int col1Based, BigDecimal value) {
        Row row = sheet.getRow(row1Based - 1);
        if (row == null) row = sheet.createRow(row1Based - 1);
        Cell cell = row.getCell(col1Based - 1);
        if (cell == null) cell = row.createCell(col1Based - 1);
        cell.setCellValue(value == null ? 0d : value.doubleValue());
    }

    private void setNumberFormat(Sheet sheet, int row1Based, int col1Based, String excelFormat) {
        Row row = sheet.getRow(row1Based - 1);
        if (row == null) row = sheet.createRow(row1Based - 1);
        Cell cell = row.getCell(col1Based - 1);
        if (cell == null) cell = row.createCell(col1Based - 1);
        var style = sheet.getWorkbook().createCellStyle();
        if (cell.getCellStyle() != null) {
            style.cloneStyleFrom(cell.getCellStyle());
        }
        DataFormat dataFormat = sheet.getWorkbook().createDataFormat();
        style.setDataFormat(dataFormat.getFormat(excelFormat));
        cell.setCellStyle(style);
    }
}
