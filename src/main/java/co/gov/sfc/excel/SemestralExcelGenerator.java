package co.gov.sfc.excel;

import co.gov.sfc.config.AiosProperties;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.CellValue;
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
import java.util.stream.Stream;

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

                CuentasData cuentas = readCuentasData(fechaCorte);
                BigDecimal aportesRecibidos = readAportesRecibidos136(fechaCorte);
                BigDecimal trm = trm(mensual);
                BigDecimal p1 = safeDivide(fondoCop, trm);

                write(hoja, 51, col, cuentas.comisiones());
                write(hoja, 52, col, cuentas.gastos());
                write(hoja, 53, col, cuentas.resultadoOperacion());
                write(hoja, 54, col, cuentas.resultadoNeto());
                write(hoja, 55, col, cuentas.admon());
                write(hoja, 56, col, cuentas.cuenta511500());
                write(hoja, 57, col, cuentas.publicidad519015());
                write(hoja, 58, col, cuentas.cuenta511500().add(cuentas.publicidad519015()));
                write(hoja, 59, col, cuentas.otros517000());
                write(hoja, 60, col, cuentas.admon().add(cuentas.otros517000()).add(cuentas.publicidad519015()));

                BigDecimal aportesUsd = safeDivide(aportesRecibidos, trm);
                BigDecimal aportantesMiles = safeDivide(mensual.aportantes(), BigDecimal.valueOf(1000));
                BigDecimal fila61 = safeDivide(aportesUsd, aportantesMiles).multiply(BigDecimal.valueOf(1000));
                write(hoja, 61, col, fila61);
                write(hoja, 62, col, safeDivide(cuentas.gastos(), aportesUsd).multiply(BigDecimal.valueOf(100)));
                write(hoja, 63, col, safeDivide(patrimonioUsd, p1).multiply(BigDecimal.valueOf(100)));
                write(hoja, 64, col, safeDivide(patrimonioUsd, mensual.afiliados()).multiply(BigDecimal.valueOf(1_000_000)));
                write(hoja, 65, col, safeDivide(cuentas.resultadoNeto(), cuentas.comisiones()).multiply(BigDecimal.valueOf(100)));
                write(hoja, 66, col, safeDivide(cuentas.resultadoNeto(), patrimonioUsd).multiply(BigDecimal.valueOf(100)));
                write(hoja, 67, col, safeDivide(cuentas.gastos(), mensual.afiliados()).multiply(BigDecimal.valueOf(1_000_000)));
                write(hoja, 68, col, safeDivide(cuentas.comisiones(), mensual.aportantes()).multiply(BigDecimal.valueOf(1_000_000)));
                write(hoja, 69, col, safeDivide(cuentas.admon(), fila61));
                write(hoja, 70, col, BigDecimal.valueOf(16));
                write(hoja, 77, col, cuentas.comisiones());
                write(hoja, 78, col, p1);
                write(hoja, 79, col, safeDivide(cuentas.comisiones(), p1));
                write(hoja, 80, col, BigDecimal.valueOf(fechaCorte.getYear() - 1994L));

                log.info("Semestral traza filas51-80: comisiones={} gastos={} resultadoOper={} resultadoNeto={} admon={} cta511500={} publicidad={} otros={} aportesRecibidosCOP={} aportesUsd={} aportantes={} fila61={} p1={}",
                        cuentas.comisiones(), cuentas.gastos(), cuentas.resultadoOperacion(), cuentas.resultadoNeto(), cuentas.admon(),
                        cuentas.cuenta511500(), cuentas.publicidad519015(), cuentas.otros517000(),
                        aportesRecibidos, aportesUsd, mensual.aportantes(), fila61, p1);
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

    private CuentasData readCuentasData(LocalDate fechaCorte) {
        Path plantilla = findPlantillaAiosFile(fechaCorte);
        try (Workbook wb = WorkbookFactory.create(plantilla.toFile(), null, true)) {
            Sheet cuentas = getSheetIgnoreCase(wb, "CUENTAS");
            if (cuentas == null) return CuentasData.ZERO;
            return new CuentasData(
                    num(cuentas, "E13"),
                    num(cuentas, "G15"),
                    num(cuentas, "E41"),
                    num(cuentas, "E44"),
                    num(cuentas, "H24"),
                    num(cuentas, "C21"),
                    num(cuentas, "C42"),
                    num(cuentas, "C37")
            );
        } catch (Exception e) {
            log.warn("No fue posible leer CUENTAS para semestral: {}", e.getMessage());
            return CuentasData.ZERO;
        }
    }

    private BigDecimal readAportesRecibidos136(LocalDate fechaCorte) {
        Path formato136 = findFormato136File(fechaCorte);
        try (Workbook wb = WorkbookFactory.create(formato136.toFile(), null, true)) {
            Sheet sheet = getSheetIgnoreCase(wb, "FORMATO OBL");
            if (sheet == null) sheet = wb.getSheetAt(0);
            Cell d6 = cell(sheet, "D6");
            d6.setCellValue(java.sql.Date.valueOf(fechaCorte));
            FormulaEvaluator evaluator = wb.getCreationHelper().createFormulaEvaluator();
            evaluator.clearAllCachedResultValues();
            BigDecimal value = num(sheet, "G6", evaluator);
            log.info("Semestral: Formato136 G6 (aportes recibidos COP)={} para fecha={}", value, fechaCorte);
            return value;
        } catch (Exception e) {
            log.warn("No fue posible leer aportes recibidos desde Formato_136_Meses: {}", e.getMessage());
            return BigDecimal.ZERO;
        }
    }

    private Path findFormato136File(LocalDate fechaCorte) {
        try (Stream<Path> paths = Files.walk(properties.insumosDir(), 3)) {
            return paths
                    .filter(Files::isRegularFile)
                    .filter(p -> p.getFileName().toString().toLowerCase(Locale.ROOT).contains("136"))
                    .findFirst()
                    .orElse(Path.of("insumos_ejemplo", "Formato_136_Meses.xlsm"));
        } catch (Exception ignore) {
            return Path.of("insumos_ejemplo", "Formato_136_Meses.xlsm");
        }
    }

    private Path findPlantillaAiosFile(LocalDate fechaCorte) {
        Path repoPath = Path.of("plantillas", "Plantilla AIOS-probable.xlsm");
        if (Files.isRegularFile(repoPath)) return repoPath;
        Path base = properties.insumosDir();
        try (Stream<Path> paths = Files.walk(base, 4)) {
            return paths
                    .filter(Files::isRegularFile)
                    .filter(p -> p.getFileName().toString().toLowerCase(Locale.ROOT).contains("plantilla aios"))
                    .findFirst()
                    .orElse(repoPath);
        } catch (Exception ignore) {
            return repoPath;
        }
    }

    private Sheet getSheetIgnoreCase(Workbook wb, String name) {
        for (int i = 0; i < wb.getNumberOfSheets(); i++) {
            Sheet sheet = wb.getSheetAt(i);
            if (sheet.getSheetName().equalsIgnoreCase(name)) {
                return sheet;
            }
        }
        return null;
    }

    private BigDecimal num(Sheet sheet, String ref) {
        return num(sheet, ref, null);
    }

    private BigDecimal num(Sheet sheet, String ref, FormulaEvaluator evaluator) {
        Cell c = cell(sheet, ref);
        if (c == null) return BigDecimal.ZERO;
        try {
            if (evaluator != null && c.getCellType() == org.apache.poi.ss.usermodel.CellType.FORMULA) {
                CellValue ev = evaluator.evaluate(c);
                if (ev != null && ev.getCellType() == org.apache.poi.ss.usermodel.CellType.NUMERIC) return BigDecimal.valueOf(ev.getNumberValue());
            }
            return switch (c.getCellType()) {
                case NUMERIC -> BigDecimal.valueOf(c.getNumericCellValue());
                case FORMULA -> {
                    if (c.getCachedFormulaResultType() == org.apache.poi.ss.usermodel.CellType.NUMERIC) {
                        yield BigDecimal.valueOf(c.getNumericCellValue());
                    }
                    yield BigDecimal.ZERO;
                }
                default -> BigDecimal.ZERO;
            };
        } catch (Exception e) {
            return BigDecimal.ZERO;
        }
    }

    private Cell cell(Sheet sheet, String ref) {
        var cr = new org.apache.poi.ss.util.CellReference(ref);
        Row row = sheet.getRow(cr.getRow());
        if (row == null) row = sheet.createRow(cr.getRow());
        Cell cell = row.getCell(cr.getCol());
        if (cell == null) cell = row.createCell(cr.getCol());
        return cell;
    }

    private record CuentasData(
            BigDecimal comisiones,
            BigDecimal gastos,
            BigDecimal resultadoOperacion,
            BigDecimal resultadoNeto,
            BigDecimal admon,
            BigDecimal cuenta511500,
            BigDecimal publicidad519015,
            BigDecimal otros517000
    ) {
        static final CuentasData ZERO = new CuentasData(
                BigDecimal.ZERO, BigDecimal.ZERO, BigDecimal.ZERO, BigDecimal.ZERO,
                BigDecimal.ZERO, BigDecimal.ZERO, BigDecimal.ZERO, BigDecimal.ZERO
        );
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
