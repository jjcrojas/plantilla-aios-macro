package co.gov.sfc.excel;

import co.gov.sfc.config.AiosProperties;
import co.gov.sfc.insumos.InsumosLocator;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
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
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.util.HashSet;
import java.util.Locale;
import java.util.Set;
import java.util.stream.Stream;

@Component
public class SemestralExcelGenerator {

    private static final Logger log = LoggerFactory.getLogger(SemestralExcelGenerator.class);

    private final AiosProperties properties;
    private final InsumosLocator locator;
    private final RentabilidadService rentabilidadService;

    public SemestralExcelGenerator(AiosProperties properties, InsumosLocator locator, RentabilidadService rentabilidadService) {
        this.properties = properties;
        this.locator = locator;
        this.rentabilidadService = rentabilidadService;
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
                BigDecimal p1 = safeDivide(mensual.vrFondo(), trm);

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
                BigDecimal patrimonioBaseMesMMCop = readPatrimonioBaseMesMMCop(fechaCorte);
                BigDecimal patrimonioBaseMesMMUsd = safeDivide(patrimonioBaseMesMMCop, trm);
                BigDecimal fila63 = safeDivide(patrimonioBaseMesMMUsd, fondoUsdMM).multiply(BigDecimal.valueOf(100));
                write(hoja, 63, col, fila63);
                write(hoja, 64, col, safeDivide(patrimonioUsd, mensual.afiliados()).multiply(BigDecimal.valueOf(1_000_000)));
                write(hoja, 65, col, safeDivide(cuentas.resultadoNeto(), cuentas.comisiones()).multiply(BigDecimal.valueOf(100)));
                write(hoja, 66, col, safeDivide(cuentas.resultadoNeto(), patrimonioUsd).multiply(BigDecimal.valueOf(100)));
                write(hoja, 67, col, safeDivide(cuentas.gastos(), mensual.afiliados()).multiply(BigDecimal.valueOf(1_000_000)));
                write(hoja, 68, col, safeDivide(cuentas.comisiones(), mensual.aportantes()).multiply(BigDecimal.valueOf(1_000_000)));
                write(hoja, 69, col, safeDivide(cuentas.admon(), fila61));
                write(hoja, 70, col, BigDecimal.valueOf(16));
                write(hoja, 77, col, cuentas.comisiones());
                // Requerimiento funcional: la fila 78 debe usar el mismo valor calculado para la fila 28.
                write(hoja, 78, col, fondoUsdMM);
                // Requerimiento funcional: fila 79 = fila 77 / fila 78.
                write(hoja, 79, col, safeDivide(cuentas.comisiones(), fondoUsdMM));
                write(hoja, 80, col, BigDecimal.valueOf(fechaCorte.getYear() - 1994L));

                log.info("Semestral traza filas51-80: comisiones={} gastos={} resultadoOper={} resultadoNeto={} admon={} cta511500={} publicidad={} otros={} aportesRecibidosCOP={} aportesUsd={} aportantes={} fila61={} p1={} fila63(%)={} patrimonioBaseMesMMCop={} patrimonioBaseMesMMUsd={} fondoUsdMM={}",
                        cuentas.comisiones(), cuentas.gastos(), cuentas.resultadoOperacion(), cuentas.resultadoNeto(), cuentas.admon(),
                        cuentas.cuenta511500(), cuentas.publicidad519015(), cuentas.otros517000(),
                        aportesRecibidos, aportesUsd, mensual.aportantes(), fila61, p1, fila63, patrimonioBaseMesMMCop, patrimonioBaseMesMMUsd, fondoUsdMM);
                BigDecimal comisionPromedioPct = promedioComisionObligatoria(trimestral).multiply(BigDecimal.valueOf(100));
                write(hoja, 71, col, comisionPromedioPct);
                write(hoja, 72, col, BigDecimal.ZERO);
                write(hoja, 73, col, BigDecimal.ZERO);
                BigDecimal aporteTrabajador = BigDecimal.valueOf(3).subtract(comisionPromedioPct).multiply(BigDecimal.valueOf(0.25));
                BigDecimal aporteEmpleador = BigDecimal.valueOf(3).subtract(comisionPromedioPct).multiply(BigDecimal.valueOf(0.75));
                write(hoja, 74, col, aporteTrabajador);
                write(hoja, 75, col, aporteEmpleador);
                write(hoja, 76, col, BigDecimal.ZERO);
                log.info("Semestral traza filas71-76: comisionPromedioPct={} aporteTrabajador={} aporteEmpleador={}",
                        comisionPromedioPct, aporteTrabajador, aporteEmpleador);
                Rentabilidades rent = readRentabilidades(fechaCorte);
                write(hoja, 82, col, rent.nominal10());
                write(hoja, 83, col, rent.real10());
                write(hoja, 84, col, rent.nominal5());
                write(hoja, 85, col, rent.real5());
                write(hoja, 86, col, rent.nominal3());
                write(hoja, 87, col, rent.real3());
                write(hoja, 88, col, rent.nominal1());
                write(hoja, 89, col, rent.real1());
                setNumberFormat(hoja, 82, col, "#,##0.00%");
                setNumberFormat(hoja, 83, col, "#,##0.00%");
                setNumberFormat(hoja, 84, col, "#,##0.00%");
                setNumberFormat(hoja, 85, col, "#,##0.00%");
                setNumberFormat(hoja, 86, col, "#,##0.00%");
                setNumberFormat(hoja, 87, col, "#,##0.00%");
                setNumberFormat(hoja, 88, col, "#,##0.00%");
                setNumberFormat(hoja, 89, col, "#,##0.00%");
                log.info("Semestral traza rentabilidades: 10y(nom={},real={}) 5y(nom={},real={}) 3y(nom={},real={}) 1y(nom={},real={})",
                        rent.nominal10(), rent.real10(), rent.nominal5(), rent.real5(), rent.nominal3(), rent.real3(), rent.nominal1(), rent.real1());

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

    private Rentabilidades readRentabilidades(LocalDate fechaCorte) {
        Path rentFile = findRentModeradoFile(fechaCorte);
        Path valoresModerado = findValoresFondoModerFile(fechaCorte);
        var y10 = calcularRentabilidadPorHorizonte(valoresModerado, rentFile, fechaCorte, 10);
        var y5 = calcularRentabilidadPorHorizonte(valoresModerado, rentFile, fechaCorte, 5);
        var y3 = calcularRentabilidadPorHorizonte(valoresModerado, rentFile, fechaCorte, 3);
        var y1 = calcularRentabilidadPorHorizonte(valoresModerado, rentFile, fechaCorte, 1);
        return new Rentabilidades(y10.nominal(), y10.real(), y5.nominal(), y5.real(), y3.nominal(), y3.real(), y1.nominal(), y1.real());
    }

    private RentPair calcularRentabilidadPorHorizonte(
            Path valoresModerado,
            Path rentFile,
            LocalDate fechaCorte,
            int anios
    ) {
        var r = rentabilidadService.calcularRentabilidad(valoresModerado, rentFile, fechaCorte, anios);
        log.info("Rent semestral {}y (NAV+IPC): ini={} fin={} nominal={} real={} valoresFile={} rentFile={}",
                anios, r.fechaInicio(), r.fechaFin(), r.rentabilidadNominal(), r.rentabilidadReal(),
                valoresModerado.toAbsolutePath(), rentFile.toAbsolutePath());
        return new RentPair(r.rentabilidadNominal(), r.rentabilidadReal());
    }

    private RentPair calcularRentabilidad(Sheet consolidado, FormulaEvaluator evaluator, LocalDate fechaInicial, LocalDate fechaFinal) {
        Cell d4 = cell(consolidado, "D4");
        Cell d5 = cell(consolidado, "D5");
        d4.setCellValue(java.sql.Date.valueOf(fechaInicial));
        d5.setCellValue(java.sql.Date.valueOf(fechaFinal));
        evaluator.clearAllCachedResultValues();
        Cell d10 = cell(consolidado, "D10");
        Cell d11 = cell(consolidado, "D11");
        BigDecimal real;
        BigDecimal nominal;
        try {
            evaluator.notifyUpdateCell(d4);
            evaluator.notifyUpdateCell(d5);
            forceRecalculateConsolidadoInputs(consolidado, evaluator);
            CellValue ev10 = evaluator.evaluate(d10);
            CellValue ev11 = evaluator.evaluate(d11);
            real = (ev10 != null && ev10.getCellType() == org.apache.poi.ss.usermodel.CellType.NUMERIC)
                    ? BigDecimal.valueOf(ev10.getNumberValue())
                    : num(consolidado, "D10");
            nominal = (ev11 != null && ev11.getCellType() == org.apache.poi.ss.usermodel.CellType.NUMERIC)
                    ? BigDecimal.valueOf(ev11.getNumberValue())
                    : num(consolidado, "D11");
            log.info("Rent moderado detalle eval: ini={} fin={} D10[type={},cached={},eval={}] D11[type={},cached={},eval={}]",
                    fechaInicial, fechaFinal,
                    d10.getCellType(), num(consolidado, "D10"), formatCellValue(ev10),
                    d11.getCellType(), num(consolidado, "D11"), formatCellValue(ev11));
        } catch (Exception e) {
            real = num(consolidado, "D10");
            nominal = num(consolidado, "D11");
            log.warn("Rent moderado: evaluator falló para inicio={} fin={}; se usan valores cacheados D10/D11. Causa={}",
                    fechaInicial, fechaFinal, e.getMessage());
        }
        log.info("Rent moderado (solo D4/D5->D10/D11): D4(inicio)={} D5(fin)={} => D11 nominal={} D10 real={}",
                fechaInicial, fechaFinal, nominal, real);
        return new RentPair(nominal, real);
    }

    private void forceRecalculateConsolidadoInputs(Sheet consolidado, FormulaEvaluator evaluator) {
        int maxRow = Math.min(consolidado.getLastRowNum(), 20);
        for (int r = 0; r <= maxRow; r++) {
            Row row = consolidado.getRow(r);
            if (row == null) continue;
            int lastCell = Math.min(Math.max(row.getLastCellNum(), (short) 1), 20);
            for (int c = 0; c < lastCell; c++) {
                Cell cell = row.getCell(c);
                if (cell == null) continue;
                if (cell.getCellType() == CellType.FORMULA) {
                    try {
                        evaluator.evaluateFormulaCell(cell);
                    } catch (Exception ignored) {
                        // Celdas con dependencias externas pueden fallar; D10/D11 se intentan evaluar al final.
                    }
                }
            }
        }
    }

    private String formatCellValue(CellValue cellValue) {
        if (cellValue == null) return "null";
        if (cellValue.getCellType() == CellType.NUMERIC) {
            return BigDecimal.valueOf(cellValue.getNumberValue()).toPlainString();
        }
        if (cellValue.getCellType() == CellType.STRING) {
            return cellValue.getStringValue();
        }
        if (cellValue.getCellType() == CellType.BOOLEAN) {
            return String.valueOf(cellValue.getBooleanValue());
        }
        if (cellValue.getCellType() == CellType.ERROR) {
            return "ERROR:" + cellValue.getErrorValue();
        }
        return cellValue.formatAsString();
    }

    private RentPair leerRentabilidadDesdeSerieConsolidado(Sheet consolidado, LocalDate fechaInicial, LocalDate fechaFinal) {
        Row rowIni = consolidado.getRow(3);   // fila 4
        Row rowFin = consolidado.getRow(4);   // fila 5
        if (rowIni == null || rowFin == null) return new RentPair(BigDecimal.ZERO, BigDecimal.ZERO);
        int last = Math.max(rowIni.getLastCellNum(), rowFin.getLastCellNum());
        int fallbackColByIniOnly = -1;
        for (int c = 3; c < Math.max(last, 4); c++) { // desde columna D
            LocalDate ini = cellAsDate(rowIni.getCell(c));
            LocalDate fin = cellAsDate(rowFin.getCell(c));
            if (fechaInicial.equals(ini) && fallbackColByIniOnly < 0) {
                fallbackColByIniOnly = c;
            }
            if (fechaInicial.equals(ini) && fechaFinal.equals(fin)) {
                BigDecimal real = num(consolidado, 10, c + 1);     // fila 10
                BigDecimal nominal = num(consolidado, 11, c + 1);  // fila 11
                log.info("Rent serie consolidado match exacto: col={} ini={} fin={} nominal(row11)={} real(row10)={}",
                        c + 1, ini, fin, nominal, real);
                return new RentPair(nominal, real);
            }
        }
        if (fallbackColByIniOnly >= 0) {
            LocalDate ini = cellAsDate(rowIni.getCell(fallbackColByIniOnly));
            LocalDate fin = cellAsDate(rowFin.getCell(fallbackColByIniOnly));
            BigDecimal real = num(consolidado, 10, fallbackColByIniOnly + 1);
            BigDecimal nominal = num(consolidado, 11, fallbackColByIniOnly + 1);
            log.info("Rent serie consolidado match por fecha inicial: col={} ini={} fin={} nominal(row11)={} real(row10)={}",
                    fallbackColByIniOnly + 1, ini, fin, nominal, real);
            return new RentPair(nominal, real);
        }
        log.warn("Rent serie consolidado: no hubo match de columna para ini={} fin={}; se usará fallback D10/D11 o tabla.",
                fechaInicial, fechaFinal);
        return new RentPair(BigDecimal.ZERO, BigDecimal.ZERO);
    }

    private RentPair calcularRentabilidadDesdeTabla(Sheet consolidado, LocalDate fechaInicial, LocalDate fechaFinal) {
        BigDecimal eIni = lookupByDate(consolidado, 5, fechaInicial);
        BigDecimal eFin = lookupByDate(consolidado, 5, fechaFinal);
        BigDecimal iIni = lookupByDate(consolidado, 9, fechaInicial);
        BigDecimal iFin = lookupByDate(consolidado, 9, fechaFinal);
        double dias = Math.max(1d, fechaFinal.toEpochDay() - fechaInicial.toEpochDay());

        BigDecimal nominal = BigDecimal.ZERO;
        if (eIni.signum() != 0 && eFin.signum() != 0) {
            nominal = BigDecimal.valueOf(Math.pow(eFin.doubleValue() / eIni.doubleValue(), 365d / dias) - 1d);
        }
        BigDecimal real = BigDecimal.ZERO;
        if (iIni.signum() != 0 && iFin.signum() != 0) {
            real = BigDecimal.valueOf(Math.pow(iFin.doubleValue() / iIni.doubleValue(), 365d / dias) - 1d);
        }
        return new RentPair(nominal, real);
    }

    private BigDecimal lookupByDate(Sheet sheet, int valueCol1Based, LocalDate target) {
        BigDecimal exacta = null;
        BigDecimal anterior = null;
        LocalDate fechaAnterior = LocalDate.MIN;
        int last = sheet.getLastRowNum() + 1;
        for (int r = 14; r <= last; r++) {
            Row row = sheet.getRow(r - 1);
            if (row == null) continue;
            LocalDate fechaFila = cellAsDate(row.getCell(0));
            if (fechaFila == null) continue;
            BigDecimal valor = num(sheet, r, valueCol1Based);
            if (fechaFila.equals(target) && valor.signum() != 0) {
                exacta = valor;
                break;
            }
            if (!fechaFila.isAfter(target) && fechaFila.isAfter(fechaAnterior) && valor.signum() != 0) {
                fechaAnterior = fechaFila;
                anterior = valor;
            }
        }
        return exacta != null ? exacta : (anterior != null ? anterior : BigDecimal.ZERO);
    }

    private LocalDate cellAsDate(Cell cell) {
        if (cell == null) return null;
        try {
            if (cell.getCellType() == org.apache.poi.ss.usermodel.CellType.NUMERIC) {
                if (org.apache.poi.ss.usermodel.DateUtil.isCellDateFormatted(cell)) {
                    return cell.getDateCellValue().toInstant().atZone(ZoneId.systemDefault()).toLocalDate();
                }
                double excel = cell.getNumericCellValue();
                if (excel > 10_000d && excel < 100_000d) {
                    return org.apache.poi.ss.usermodel.DateUtil.getJavaDate(excel).toInstant().atZone(ZoneId.systemDefault()).toLocalDate();
                }
            }
            String txt = new DataFormatter(Locale.forLanguageTag("es-CO")).formatCellValue(cell);
            if (txt == null || txt.isBlank()) return null;
            String v = txt.trim().toLowerCase(Locale.ROOT).replace(".", "");
            DateTimeFormatter[] fmts = new DateTimeFormatter[]{
                    DateTimeFormatter.ofPattern("d-MMM-yy", new Locale("es", "CO")),
                    DateTimeFormatter.ofPattern("d-MMM-yyyy", new Locale("es", "CO")),
                    DateTimeFormatter.ofPattern("d/M/yyyy"),
                    DateTimeFormatter.ofPattern("d/M/yy"),
                    DateTimeFormatter.ISO_LOCAL_DATE
            };
            for (DateTimeFormatter f : fmts) {
                try {
                    return LocalDate.parse(v, f);
                } catch (Exception ignored) {
                }
            }
            return null;
        } catch (Exception e) {
            return null;
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

    private BigDecimal readPatrimonioBaseMesMMCop(LocalDate fechaCorte) {
        Path plantilla = findPlantillaAiosFile(fechaCorte);
        try (Workbook wb = WorkbookFactory.create(plantilla.toFile(), null, true)) {
            Sheet cuentas = getSheetIgnoreCase(wb, "cuentas");
            Sheet baseMes = getSheetIgnoreCase(wb, "base mes");
            if (cuentas == null || baseMes == null) {
                log.warn("Patrimonio base mes: no se encontró hoja cuentas/base mes en {}", plantilla.toAbsolutePath());
                return BigDecimal.ZERO;
            }
            LocalDate fechaBaseMes = LocalDate.of(fechaCorte.getYear(), fechaCorte.getMonth(), 1);
            int serialFecha = (int) Math.round(org.apache.poi.ss.usermodel.DateUtil.getExcelDate(java.sql.Date.valueOf(fechaBaseMes)));
            int serialFechaCorte = (int) Math.round(org.apache.poi.ss.usermodel.DateUtil.getExcelDate(java.sql.Date.valueOf(fechaCorte)));
            String cuentaPatrimonio = "300000";
            Set<String> entidades = new HashSet<>();
            for (int r = 1; r <= 4; r++) { // J1:J4
                Row row = cuentas.getRow(r - 1);
                if (row == null) continue;
                Cell c = row.getCell(9);
                if (c == null) continue;
                String entidad = normalize(c.toString());
                if (!entidad.isBlank()) entidades.add(entidad);
            }
            if (entidades.isEmpty()) {
                log.warn("Patrimonio base mes: no se encontraron administradoras en cuentas!J1:J4");
                return BigDecimal.ZERO;
            }

            Set<String> keys = new HashSet<>();
            for (String entidad : entidades) {
                keys.add(entidad + "-" + serialFecha + "-" + cuentaPatrimonio);
            }

            BigDecimal sumaCop = BigDecimal.ZERO;
            Set<String> encontradas = new HashSet<>();
            int last = baseMes.getLastRowNum() + 1;
            for (int r = 2; r <= last; r++) {
                Row row = baseMes.getRow(r - 1);
                if (row == null) continue;
                Cell keyCell = row.getCell(0); // col A llave
                if (keyCell == null) continue;
                String key = normalize(keyCell.toString());
                if (!keys.contains(key)) continue;
                BigDecimal valor = num(baseMes, r, 6); // col F valor
                if (valor.signum() > 0) {
                    sumaCop = sumaCop.add(valor);
                    encontradas.add(key);
                    log.info("Patrimonio base mes match: key={} valorCOP={}", key, valor);
                }
                if (encontradas.size() == keys.size()) break;
            }
            BigDecimal mmCop = sumaCop.divide(BigDecimal.valueOf(1_000_000), 8, RoundingMode.HALF_UP);
            log.info("Patrimonio base mes total: fechaParametro={} fechaBaseMes={} serialBaseMes={} serialFechaCorte={} entidades={} matches={} sumaCOP={} sumaMMCOP={}",
                    fechaCorte, fechaBaseMes, serialFecha, serialFechaCorte, entidades, encontradas.size(), sumaCop, mmCop);
            if (encontradas.size() < keys.size()) {
                log.warn("Patrimonio base mes incompleto: esperadas={} encontradas={} faltantes={}",
                        keys.size(), encontradas.size(), keys.stream().filter(k -> !encontradas.contains(k)).toList());
                // Fallback defensivo: si no hay match con serial del primer día de mes, intentar serial exacto de fecha de corte.
                if (serialFechaCorte != serialFecha) {
                    Set<String> keysCorte = new HashSet<>();
                    for (String entidad : entidades) {
                        keysCorte.add(entidad + "-" + serialFechaCorte + "-" + cuentaPatrimonio);
                    }
                    BigDecimal sumaCopCorte = BigDecimal.ZERO;
                    Set<String> encontradasCorte = new HashSet<>();
                    for (int r = 2; r <= last; r++) {
                        Row row = baseMes.getRow(r - 1);
                        if (row == null) continue;
                        Cell keyCell = row.getCell(0);
                        if (keyCell == null) continue;
                        String key = normalize(keyCell.toString());
                        if (!keysCorte.contains(key)) continue;
                        BigDecimal valor = num(baseMes, r, 6);
                        if (valor.signum() > 0) {
                            sumaCopCorte = sumaCopCorte.add(valor);
                            encontradasCorte.add(key);
                            log.info("Patrimonio base mes fallback(fecha corte) match: key={} valorCOP={}", key, valor);
                        }
                        if (encontradasCorte.size() == keysCorte.size()) break;
                    }
                    if (encontradasCorte.size() > encontradas.size()) {
                        BigDecimal mmCopCorte = sumaCopCorte.divide(BigDecimal.valueOf(1_000_000), 8, RoundingMode.HALF_UP);
                        log.info("Patrimonio base mes fallback usado con serial fecha corte: serial={} matches={} sumaCOP={} sumaMMCOP={}",
                                serialFechaCorte, encontradasCorte.size(), sumaCopCorte, mmCopCorte);
                        return mmCopCorte;
                    }
                }
            }
            return mmCop;
        } catch (Exception e) {
            log.warn("No fue posible leer patrimonio desde base mes: {}", e.getMessage());
            return BigDecimal.ZERO;
        }
    }

    private Path findRentModeradoFile(LocalDate fechaCorte) {
        try {
            return locator.findRequired("Rent_Vr_Uni_Moderado", fechaCorte);
        } catch (Exception ignore) {
            return Path.of("insumos_ejemplo", "Rent_Vr_Uni_Moderado.xlsm");
        }
    }

    private Path findValoresFondoModerFile(LocalDate fechaCorte) {
        try {
            return locator.findRequired("Valores_Fondo_Moder", fechaCorte);
        } catch (Exception e1) {
            try {
                return locator.findRequired("MODERADO", fechaCorte);
            } catch (Exception e2) {
                return Path.of("insumos_ejemplo", "MODERADO Junio 2025.xls");
            }
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

    private BigDecimal num(Sheet sheet, int row1Based, int col1Based) {
        Row row = sheet.getRow(row1Based - 1);
        if (row == null) return BigDecimal.ZERO;
        Cell c = row.getCell(col1Based - 1);
        if (c == null) return BigDecimal.ZERO;
        try {
            return switch (c.getCellType()) {
                case NUMERIC -> BigDecimal.valueOf(c.getNumericCellValue());
                case FORMULA -> c.getCachedFormulaResultType() == org.apache.poi.ss.usermodel.CellType.NUMERIC
                        ? BigDecimal.valueOf(c.getNumericCellValue())
                        : BigDecimal.ZERO;
                default -> BigDecimal.ZERO;
            };
        } catch (Exception e) {
            return BigDecimal.ZERO;
        }
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

    private record RentPair(BigDecimal nominal, BigDecimal real) {}

    private record Rentabilidades(
            BigDecimal nominal10, BigDecimal real10,
            BigDecimal nominal5, BigDecimal real5,
            BigDecimal nominal3, BigDecimal real3,
            BigDecimal nominal1, BigDecimal real1
    ) {
        static final Rentabilidades ZERO = new Rentabilidades(
                BigDecimal.ZERO, BigDecimal.ZERO,
                BigDecimal.ZERO, BigDecimal.ZERO,
                BigDecimal.ZERO, BigDecimal.ZERO,
                BigDecimal.ZERO, BigDecimal.ZERO
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
