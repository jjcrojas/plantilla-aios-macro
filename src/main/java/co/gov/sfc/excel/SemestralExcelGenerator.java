package co.gov.sfc.excel;

import co.gov.sfc.config.AiosProperties;
import co.gov.sfc.insumos.InsumosLocator;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
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
import java.util.Date;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Locale;
import java.util.Map;
import java.util.Set;
import java.util.stream.Stream;

@Component
public class SemestralExcelGenerator {

    private static final Logger log = LoggerFactory.getLogger(SemestralExcelGenerator.class);

    private final AiosProperties properties;
    private final InsumosLocator locator;
    private final RentabilidadService rentabilidadService;
    private final Formato493QueryService formato493QueryService;
    private final Formato495QueryService formato495QueryService;
    private final Formato136QueryService formato136QueryService;
    private final ComisionesSemestralQueryService comisionesSemestralQueryService;

    public SemestralExcelGenerator(AiosProperties properties, InsumosLocator locator, RentabilidadService rentabilidadService, Formato493QueryService formato493QueryService, Formato495QueryService formato495QueryService, Formato136QueryService formato136QueryService, ComisionesSemestralQueryService comisionesSemestralQueryService) {
        this.properties = properties;
        this.locator = locator;
        this.rentabilidadService = rentabilidadService;
        this.formato493QueryService = formato493QueryService;
        this.formato495QueryService = formato495QueryService;
        this.formato136QueryService = formato136QueryService;
        this.comisionesSemestralQueryService = comisionesSemestralQueryService;
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
                java.util.Map<Integer, String> detallesFilas = new java.util.LinkedHashMap<>();

                // Bloque A - principales (según EscribirSemestral_Integral)
                writeFilasAfiliadosDisponibilidad(hoja, col, mensual);
                write(hoja, 4, col, pct(safeDivide(mensual.afiliadosMenor30(), mensual.afiliados())));
                write(hoja, 5, col, pct(safeDivide(mensual.afiliados30a44(), mensual.afiliados())));
                write(hoja, 6, col, pct(safeDivide(mensual.afiliados45a59(), mensual.afiliados())));
                write(hoja, 7, col, pct(safeDivide(mensual.afiliadosMayor60(), mensual.afiliados())));
                write(hoja, 8, col, BigDecimal.valueOf(100));
                write(hoja, 9, col, divide(mensual.afiliados(), BigDecimal.valueOf(1000)));
                write(hoja, 10, col, pct(safeDivide(mensual.mujeres(), mensual.afiliados())));
                write(hoja, 11, col, mensual.aportantesSemestral());
                write(hoja, 12, col, pct(safeDivide(mensual.afiliados(), mensual.pea())));
                write(hoja, 13, col, pct(safeDivide(mensual.aportantesSemestral(), mensual.pea())));
                write(hoja, 14, col, pct(safeDivide(mensual.aportantesSemestral(), mensual.afiliados())));
                write(hoja, 15, col, mensual.smColombiaUsd());
                Formato495QueryService.PensionadosResumen pensionadosSemestral = formato495QueryService.leerResumen(fechaCorte);
                BigDecimal totalPensionadosSemestral = pensionadosSemestral.total();
                PensionadosPorEntidad pensionadosPorEntidad = new PensionadosPorEntidad(
                        pensionadosSemestral.invalidez(), pensionadosSemestral.vejez(), pensionadosSemestral.sobrevivencia());
                BigDecimal fila17 = safeDivide(pensionadosPorEntidad.invalidez(), totalPensionadosSemestral);
                BigDecimal fila18 = safeDivide(pensionadosPorEntidad.vejez(), totalPensionadosSemestral);
                BigDecimal fila19 = safeDivide(pensionadosPorEntidad.sobrevivencia(), totalPensionadosSemestral);
                write(hoja, 16, col, totalPensionadosSemestral);
                write(hoja, 17, col, fila17);
                write(hoja, 18, col, fila18);
                write(hoja, 19, col, fila19);
                BigDecimal fila25 = readFila25Trimestral493(fechaCorte);
                write(hoja, 25, col, fila25);
                log.info("Semestral: fila25=fallecidos/1000 desde query Formato 493 para fechaCorte={} => {}.", fechaCorte, fila25);
                log.info("Semestral: fila16(total_pen query 495)={}, fila17(inv query 495/total)={}, fila18(vej query 495/total)={}, fila19(sob query 495/total)={} numeradores(inv={}, vej={}, sob={}) para fecha={} col={}.",
                        totalPensionadosSemestral,
                        fila17,
                        fila18,
                        fila19,
                        pensionadosPorEntidad.invalidez(),
                        pensionadosPorEntidad.vejez(),
                        pensionadosPorEntidad.sobrevivencia(),
                        fechaCorte, col);
                write(hoja, 26, col, mensual.traspasosSistema());
                write(hoja, 27, col, safeDivide(mensual.traspasosSistema(), mensual.afiliados()));
                setNumberFormat(hoja, 27, col, "#,##0.00%");
                BigDecimal fondoUsdMM = safeDivide(mensual.vrFondo(), trm(mensual));
                write(hoja, 28, col, fondoUsdMM);
                BigDecimal pibUsd = safeDivide(mensual.pibSemestral(), trm(mensual));
                BigDecimal ratioFondosPib = safeDivide(fondoUsdMM, pibUsd).multiply(BigDecimal.valueOf(100));
                write(hoja, 29, col, ratioFondosPib);
                setNumberFormat(hoja, 29, col, "#,##0.00");
                log.info("Semestral traza fila29: fondoUsdMM={} pibSemestralCOP={} trm={} pibUsd={} ratioFondosPib={} fecha={} col={}",
                        fondoUsdMM, mensual.pibSemestral(), trm(mensual), pibUsd, ratioFondosPib, fechaCorte, col);
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
                DatoDetalle fila44Pct = readFila44DesdeLimites(fechaCorte);
                write(hoja, 44, col, fila44Pct.valor());
                detallesFilas.put(44, fila44Pct.detalle());
                DatoDetalle deudaGubernamentalTotal = readDeudaGubernamentalTotal(fechaCorte);
                BigDecimal fila45 = safeDivide(fondoUsdMM, deudaGubernamentalTotal.valor());
                write(hoja, 45, col, fila45);
                setNumberFormat(hoja, 45, col, "#,##0.00%");
                detallesFilas.put(45, "operando fila28 fondoUsdMM=" + fondoUsdMM + "; operando deudaGubernamentalTotalUSD=" + deudaGubernamentalTotal.valor() + "; " + deudaGubernamentalTotal.detalle());
                write(hoja, 46, col, BigDecimal.valueOf(4));
                DatoDetalle fila47 = new DatoDetalle(
                        safeDivide(mensual.porcVrFondo(), BigDecimal.valueOf(100)),
                        "fuente=Query Teradata NEGFID_INSUMO_ENTIDAD; operación=(Protección+Porvenir)/total sistema.");
                write(hoja, 47, col, fila47.valor());
                detallesFilas.put(47, fila47.detalle());
                setNumberFormat(hoja, 47, col, "#,##0.00%");
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

                BigDecimal comisionesFila51 = comisionesSemestralQueryService.leer411500(fechaCorte, trm);
                write(hoja, 51, col, comisionesFila51);
                BigDecimal gastosFila52 = comisionesSemestralQueryService.leerGastosOperativos(fechaCorte, trm);
                write(hoja, 52, col, gastosFila52);
                write(hoja, 53, col, cuentas.resultadoOperacion());
                write(hoja, 54, col, cuentas.resultadoNeto());
                write(hoja, 55, col, cuentas.admon());
                BigDecimal fila56 = safeDivide(cuentas.cuenta511500(), trm);
                BigDecimal fila57 = safeDivide(cuentas.cuenta511527(), trm);
                BigDecimal fila58Base = cuentas.cuenta511500().add(cuentas.cuenta511527());
                BigDecimal fila58 = safeDivide(fila58Base, trm);
                BigDecimal fila59 = safeDivide(cuentas.otrosGastosOperacion(), trm);
                BigDecimal fila60 = safeDivide(cuentas.gastoOperacion510000(), trm);
                write(hoja, 56, col, fila56);
                write(hoja, 57, col, fila57);
                write(hoja, 58, col, fila58);
                write(hoja, 59, col, fila59);
                write(hoja, 60, col, fila60);
                detallesFilas.put(56, "detalle fuente cuenta 511500: hoja=cuentas celda=C21 valorCOP=" + cuentas.cuenta511500() + "; TRM=" + trm + "; operación=C21/TRM.");
                detallesFilas.put(57, "detalle fuente cuenta 511527: hoja=cuentas celda=C22 valorCOP=" + cuentas.cuenta511527() + "; TRM=" + trm + "; operación=C22/TRM.");
                detallesFilas.put(58, "detalle fuente cuentas 511500+511527: hoja=cuentas celdas C21=" + cuentas.cuenta511500() + ", C22=" + cuentas.cuenta511527() + "; sumaCOP=" + fila58Base + "; TRM=" + trm + "; operación=(C21+C22)/TRM.");
                detallesFilas.put(59, "detalle fuente otros gastos de operación: hoja=cuentas celdas C24,C28,C29,C31,C32,C33,C34,C35,C36,C37,C38 sumaCOP=" + cuentas.otrosGastosOperacion() + "; TRM=" + trm + "; operación=suma/TRM.");
                detallesFilas.put(60, "detalle fuente cuenta 510000: hoja=cuentas celda=C15 valorCOP=" + cuentas.gastoOperacion510000() + "; TRM=" + trm + "; operación=C15/TRM.");

                BigDecimal aportesUsd = safeDivide(aportesRecibidos, trm);
                BigDecimal aportantesMiles = safeDivide(mensual.aportantesSemestral(), BigDecimal.valueOf(1000));
                BigDecimal fila61 = safeDivide(aportesUsd, aportantesMiles).multiply(BigDecimal.valueOf(1000));
                write(hoja, 61, col, fila61);
                write(hoja, 62, col, safeDivide(gastosFila52, aportesUsd).multiply(BigDecimal.valueOf(100)));
                BigDecimal patrimonioBaseMesMMCop = readPatrimonioBaseMesMMCop(fechaCorte);
                BigDecimal patrimonioBaseMesMMUsd = safeDivide(patrimonioBaseMesMMCop, trm);
                BigDecimal fila63 = safeDivide(patrimonioBaseMesMMUsd, fondoUsdMM).multiply(BigDecimal.valueOf(100));
                write(hoja, 63, col, fila63);
                write(hoja, 64, col, safeDivide(patrimonioUsd, mensual.afiliados()).multiply(BigDecimal.valueOf(1_000_000)));
                write(hoja, 65, col, safeDivide(cuentas.resultadoNeto(), comisionesFila51).multiply(BigDecimal.valueOf(100)));
                write(hoja, 66, col, safeDivide(cuentas.resultadoNeto(), patrimonioUsd).multiply(BigDecimal.valueOf(100)));
                write(hoja, 67, col, safeDivide(gastosFila52, mensual.afiliados()).multiply(BigDecimal.valueOf(1_000_000)));
                write(hoja, 68, col, safeDivide(comisionesFila51, mensual.aportantesSemestral()).multiply(BigDecimal.valueOf(1_000_000)));
                write(hoja, 69, col, safeDivide(cuentas.admon(), fila61));
                write(hoja, 70, col, BigDecimal.valueOf(16));
                write(hoja, 77, col, comisionesFila51);
                // Requerimiento funcional: la fila 78 debe usar el mismo valor calculado para la fila 28.
                write(hoja, 78, col, fondoUsdMM);
                // Requerimiento funcional: fila 79 = fila 77 / fila 78.
                write(hoja, 79, col, safeDivide(comisionesFila51, fondoUsdMM));
                write(hoja, 80, col, BigDecimal.valueOf(fechaCorte.getYear() - 1994L));

                log.info("Semestral traza filas51-80: comisiones={} gastos={} resultadoOper={} resultadoNeto={} admon={} fila56(511500/TRM)={} fila57(511527/TRM)={} fila58((511500+511527)/TRM)={} fila59(otros/TRM)={} fila60(510000/TRM)={} aportesRecibidosCOP={} aportesUsd={} aportantes={} fila61={} p1={} fila63(%)={} patrimonioBaseMesMMCop={} patrimonioBaseMesMMUsd={} fondoUsdMM={}",
                        comisionesFila51, gastosFila52, cuentas.resultadoOperacion(), cuentas.resultadoNeto(), cuentas.admon(),
                        fila56, fila57, fila58, fila59, fila60,
                        aportesRecibidos, aportesUsd, mensual.aportantesSemestral(), fila61, p1, fila63, patrimonioBaseMesMMCop, patrimonioBaseMesMMUsd, fondoUsdMM);
                BigDecimal comisionPromedioPct = promedioComisionObligatoria(trimestral);
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
                write(hoja, 82, col, rent.nominal10().multiply(BigDecimal.valueOf(100)));
                write(hoja, 83, col, rent.real10().multiply(BigDecimal.valueOf(100)));
                write(hoja, 84, col, rent.nominal5().multiply(BigDecimal.valueOf(100)));
                write(hoja, 85, col, rent.real5().multiply(BigDecimal.valueOf(100)));
                write(hoja, 86, col, rent.nominal3().multiply(BigDecimal.valueOf(100)));
                write(hoja, 87, col, rent.real3().multiply(BigDecimal.valueOf(100)));
                write(hoja, 88, col, rent.nominal1().multiply(BigDecimal.valueOf(100)));
                write(hoja, 89, col, rent.real1().multiply(BigDecimal.valueOf(100)));
                setNumberFormat(hoja, 82, col, "#,##0.00");
                setNumberFormat(hoja, 83, col, "#,##0.00");
                setNumberFormat(hoja, 84, col, "#,##0.00");
                setNumberFormat(hoja, 85, col, "#,##0.00");
                setNumberFormat(hoja, 86, col, "#,##0.00");
                setNumberFormat(hoja, 87, col, "#,##0.00");
                setNumberFormat(hoja, 88, col, "#,##0.00");
                setNumberFormat(hoja, 89, col, "#,##0.00");
                log.info("Semestral traza rentabilidades: 10y(nom={},real={}) 5y(nom={},real={}) 3y(nom={},real={}) 1y(nom={},real={})",
                        rent.nominal10().multiply(BigDecimal.valueOf(100)), rent.real10().multiply(BigDecimal.valueOf(100)),
                        rent.nominal5().multiply(BigDecimal.valueOf(100)), rent.real5().multiply(BigDecimal.valueOf(100)),
                        rent.nominal3().multiply(BigDecimal.valueOf(100)), rent.real3().multiply(BigDecimal.valueOf(100)),
                        rent.nominal1().multiply(BigDecimal.valueOf(100)), rent.real1().multiply(BigDecimal.valueOf(100)));
                logFilasSemestral(hoja, col, fechaCorte, mensual, trimestral, detallesFilas);
                normalizarEstilosSemestral(hoja);

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


    private void logFilasSemestral(Sheet hoja, int col, LocalDate fechaCorte, MensualData mensual, TrimestralData trimestral, java.util.Map<Integer, String> detallesFilas) {
        String limites = rutaInsumo("LIMITES", () -> locator.findRequired("LIMITES", fechaCorte));
        String pibPeaTrmDg = rutaInsumo("PIB_PEA_TRM_DG", () -> locator.findRequired("PIB_PEA_TRM_DG", fechaCorte));
        String queryAportes136 = "Query Teradata prod_dwh_consulta.negfid_insumo_entidad (nivel1=136,nivel2=2,nivel3=4,nivel4=10)";
        String plantillaAios = rutaInsumo("Plantilla AIOS-probable", () -> findPlantillaAiosFile(fechaCorte));
        String rentVrUni = rutaInsumo("Rent_Vr_Uni_Moderado", () -> findRentModeradoFile(fechaCorte));
        String valoresFondo = rutaInsumo("Valores_Fondo_Moder/MODERADO", () -> findValoresFondoModerFile(fechaCorte));

        java.util.Map<Integer, String> explicaciones = new java.util.LinkedHashMap<>();
        explicaciones.put(3, "valor = mensual.afiliados(); mismo total de afiliados usado en la columna B del archivo mensual; fuente=Query Teradata PROD_DWH_CONSULTA.FORMATO491 (RENGLON=999, SUM(TOTAL_AFILIADOS_TOTAL), fondos 1000/5000/6000/7000/8000).");
        explicaciones.put(4, "valor = (mensual.afiliadosMenor30() / mensual.afiliados()) * 100; afiliadosMenor30=Query Teradata (regla subcuenta/unidad captura)=" + mensual.afiliadosMenor30() + ", afiliadosTotalQuery=" + mensual.afiliados() + ".");
        explicaciones.put(5, "valor = (mensual.afiliados30a44() / mensual.afiliados()) * 100; afiliados30a44=Query Teradata (regla subcuenta/unidad captura)=" + mensual.afiliados30a44() + ", afiliadosTotalQuery=" + mensual.afiliados() + ".");
        explicaciones.put(6, "valor = (mensual.afiliados45a59() / mensual.afiliados()) * 100; afiliados45a59=Query Teradata (regla subcuenta/unidad captura)=" + mensual.afiliados45a59() + ", afiliadosTotalQuery=" + mensual.afiliados() + ".");
        explicaciones.put(7, "valor = (mensual.afiliadosMayor60() / mensual.afiliados()) * 100; afiliadosMayor60=Query Teradata (regla subcuenta/unidad captura)=" + mensual.afiliadosMayor60() + ", afiliadosTotalQuery=" + mensual.afiliados() + ".");
        explicaciones.put(8, "valor fijo = 100; representa el total porcentual de rangos de edad.");
        explicaciones.put(9, "valor = mensual.afiliados() / 1000; afiliadosTotalQuery=" + mensual.afiliados() + ".");
        explicaciones.put(10, "valor = (mensual.mujeres() / mensual.afiliados()) * 100; mujeres=Query Teradata PROD_DWH_CONSULTA.FORMATO491 (RENGLON=999, SUM(TOTAL_AFILIADOS_M), fondos 1000/5000/6000/7000/8000)=" + mensual.mujeres() + ", afiliadosTotalQuery=" + mensual.afiliados() + ".");
        explicaciones.put(11, "valor = mensual.aportantesSemestral(); fuente=Query Teradata PROD_DWH_CONSULTA.FORMATO491 (RENGLON=999, SUM(TOTAL_AFILIADOS_COTIZANTES), fondos 1000/5000/6000/7000/8000), aportantes=" + mensual.aportantesSemestral() + ".");
        explicaciones.put(12, "valor = (mensual.afiliados() / mensual.pea()) * 100; afiliados por query Teradata Formato491; PEA del archivo PIB_PEA_TRM_DG ruta=" + pibPeaTrmDg + ".");
        explicaciones.put(13, "valor = (mensual.aportantesSemestral() / mensual.pea()) * 100; aportantes por query Teradata sin filtro de CODIGO_ENTIDAD y PEA del archivo PIB_PEA_TRM_DG ruta=" + pibPeaTrmDg + ".");
        explicaciones.put(14, "valor = (mensual.aportantesSemestral() / mensual.afiliados()) * 100; aportantes por query Teradata sin filtro de CODIGO_ENTIDAD y afiliados por query Teradata.");
        explicaciones.put(15, "valor = salario mínimo ponderado COP calculado por query Teradata Formato491 con salario oficial desde SalarioMinimo.csv / TRM; salarioMinimoPonderadoCop=" + smCop(mensual) + "; TRM desde PIB_PEA_TRM_DG ruta=" + pibPeaTrmDg + ".");
        explicaciones.put(16, "valor = total pensionados por query Teradata PROD_DWH_CONSULTA.S9_FORMATO_495 con FECHA_CORTE, UNIDAD_CAPTURA=1 y RENGLON=200.");
        explicaciones.put(17, "valor = pensionados por invalidez por query Teradata PROD_DWH_CONSULTA.S9_FORMATO_495 / fila 16.");
        explicaciones.put(18, "valor = pensionados por vejez por query Teradata PROD_DWH_CONSULTA.S9_FORMATO_495 / fila 16.");
        explicaciones.put(19, "valor = pensionados por sobrevivencia por query Teradata PROD_DWH_CONSULTA.S9_FORMATO_495 / fila 16.");
        explicaciones.put(25, "valor = query Teradata PROD_DWH_CONSULTA.S9_FORMATO_493 fallecidos sistema / 1000; ventana de 12 meses por FECHA_CORTE, UNIDAD_CAPTURA=1 y RENGLON IN (165,170,175).");
        explicaciones.put(26, "valor = mensual.traspasosSistema(); total de traspasos del sistema leído por query Teradata PROD_DWH_CONSULTA.S9_FORMATO_493 en MensualDataReader.");
        explicaciones.put(27, "valor = mensual.traspasosSistema() / mensual.afiliados(); traspasos por query Teradata Formato493 dividido entre afiliados por query Teradata Formato491.");
        explicaciones.put(28, "valor = mensual.vrFondo() / mensual.trm(); vrFondo proviene de Query Teradata NEGFID_INSUMO_ENTIDAD, niveles 136/2/4/305, SUM(valor)/1,000,000; TRM de PIB_PEA_TRM_DG ruta=" + pibPeaTrmDg + ".");
        explicaciones.put(29, "valor = fila 28 / (mensual.pibSemestral() / mensual.trm()) * 100; se guarda en puntos porcentuales y se muestra sin símbolo %; PIB semestral y TRM desde PIB_PEA_TRM_DG ruta=" + pibPeaTrmDg + ".");
        explicaciones.put(30, "valor = mensual.total1() / mensual.trm(); total1 desde límites/composición leída por MensualDataReader; TRM desde PIB_PEA_TRM_DG ruta=" + pibPeaTrmDg + ".");
        explicaciones.put(31, "valor = mensual.dudaG(); dato de límites de inversión leído por MensualDataReader desde LIMITES/AIOS u origen equivalente, ruta=" + limites + ".");
        explicaciones.put(32, "valor = mensual.dudaEf(); dato de límites de inversión leído desde LIMITES/AIOS, ruta=" + limites + ".");
        explicaciones.put(33, "valor = mensual.dudaNf(); dato de límites de inversión leído desde LIMITES/AIOS, ruta=" + limites + ".");
        explicaciones.put(34, "valor = mensual.dudaAc(); dato de límites de inversión leído desde LIMITES/AIOS, ruta=" + limites + ".");
        explicaciones.put(35, "valor = mensual.dudaF(); dato de límites de inversión leído desde LIMITES/AIOS, ruta=" + limites + ".");
        explicaciones.put(36, "valor fijo = 0; no usa insumo externo.");
        explicaciones.put(37, "valor = mensual.dudaGe(); límite exterior leído desde LIMITES/AIOS, ruta=" + limites + ".");
        explicaciones.put(38, "valor = mensual.dudaEfe(); límite exterior leído desde LIMITES/AIOS, ruta=" + limites + ".");
        explicaciones.put(39, "valor = mensual.dudaNfe(); límite exterior leído desde LIMITES/AIOS, ruta=" + limites + ".");
        explicaciones.put(40, "valor = mensual.dudaAce(); límite exterior leído desde LIMITES/AIOS, ruta=" + limites + ".");
        explicaciones.put(41, "valor = mensual.dudaFe(); límite exterior leído desde LIMITES/AIOS, ruta=" + limites + ".");
        explicaciones.put(42, "valor fijo = 2; no usa insumo externo.");
        explicaciones.put(43, "valor = mensual.otros(); otros conceptos calculados por MensualDataReader desde los insumos mensuales.");
        explicaciones.put(44, "valor = (LIMITES hoja AIOS celdas O4 + Q4 + S4 + U4 + W4 + Y4) * 100; ruta=" + limites + ".");
        explicaciones.put(45, "valor = fila 28 / deuda gubernamental total en USD; deuda gubernamental total proviene de PIB_PEA_TRM_DG hoja Hoja1 columna M para la fecha de corte, ruta=" + pibPeaTrmDg + ".");
        explicaciones.put(46, "valor fijo = 4; no usa insumo externo.");
        explicaciones.put(47, "valor = concentración de Protección y Porvenir consultada en Teradata NEGFID_INSUMO_ENTIDAD para niveles 136/2/4/305 y la fecha de corte.");
        explicaciones.put(48, "valor = mensual.activosCuentas() / mensual.trm(); activosCuentas desde Plantilla AIOS/CUENTAS o fuente contable, ruta=" + plantillaAios + "; TRM ruta=" + pibPeaTrmDg + ".");
        explicaciones.put(49, "valor = mensual.pasivosCuentas() / mensual.trm(); pasivosCuentas desde Plantilla AIOS/CUENTAS ruta=" + plantillaAios + "; TRM ruta=" + pibPeaTrmDg + ".");
        explicaciones.put(50, "valor = (mensual.activosCuentas() - mensual.pasivosCuentas()) / mensual.trm(); activos y pasivos desde CUENTAS ruta=" + plantillaAios + "; TRM ruta=" + pibPeaTrmDg + ".");
        explicaciones.put(51, "valor = query Teradata sobre ESTFIN_INDIV, cuenta 411500: (saldo corte actual + saldo cierre anterior - saldo mismo corte anterior) / 1,000,000 / TRM; la ruta de CUENTAS ya no se usa para esta fila. Plantilla=" + plantillaAios + ".");
        explicaciones.put(52, "valor = query Teradata de gastos operativos: cuenta 510000 menos cuentas 510300, 510400, 510600, 510700, 510800, 512500, 512800, 512900 y 513900; aplica saldo corte + cierre anterior - mismo corte anterior, dividido entre 1,000,000 y TRM. Plantilla=" + plantillaAios + ".");
        explicaciones.put(53, "valor = cuentas.resultadoOperacion(); leído desde Plantilla AIOS hoja CUENTAS, ruta=" + plantillaAios + ".");
        explicaciones.put(54, "valor = cuentas.resultadoNeto(); leído desde Plantilla AIOS hoja CUENTAS, ruta=" + plantillaAios + ".");
        explicaciones.put(55, "valor = cuentas.admon(); gastos de administración desde Plantilla AIOS hoja CUENTAS, ruta=" + plantillaAios + ".");
        explicaciones.put(56, "valor = cuenta 511500 / TRM; cuenta 511500 proviene de Plantilla AIOS-probable hoja cuentas celda C21, ruta=" + plantillaAios + "; TRM ruta=" + pibPeaTrmDg + ".");
        explicaciones.put(57, "valor = cuenta 511527 / TRM; cuenta 511527 proviene de Plantilla AIOS-probable hoja cuentas celda C22, ruta=" + plantillaAios + "; TRM ruta=" + pibPeaTrmDg + ".");
        explicaciones.put(58, "valor = (cuenta 511500 + cuenta 511527) / TRM; cuentas desde Plantilla AIOS-probable hoja cuentas celdas C21 y C22, ruta=" + plantillaAios + "; TRM ruta=" + pibPeaTrmDg + ".");
        explicaciones.put(59, "valor = suma de cuentas 512000, 513000, 513500, 514000, 514500, 515000, 515500, 516000, 516500, 517000 y 517200 / TRM; celdas C24,C28,C29,C31,C32,C33,C34,C35,C36,C37,C38 de Plantilla AIOS-probable hoja cuentas, ruta=" + plantillaAios + "; TRM ruta=" + pibPeaTrmDg + ".");
        explicaciones.put(60, "valor = cuenta 510000 / TRM; cuenta 510000 proviene de Plantilla AIOS-probable hoja cuentas celda C15, ruta=" + plantillaAios + "; TRM ruta=" + pibPeaTrmDg + ".");
        explicaciones.put(61, "valor = (aportesRecibidos136 / TRM) / (mensual.aportantesSemestral() / 1000) * 1000; aportesRecibidos136 desde " + queryAportes136 + ", con b.fecha entre el día 1 del mismo mes un año antes del corte y la fecha de corte; TRM ruta=" + pibPeaTrmDg + ".");
        explicaciones.put(62, "valor = fila 52 / (aportesRecibidos136 / TRM) * 100; gastos desde query Teradata de la fila 52; plantilla=" + plantillaAios + "; aportes desde " + queryAportes136 + ".");
        explicaciones.put(63, "valor = (patrimonioBaseMesMMCop / TRM) / fila 28 * 100; patrimonio base_mes desde Plantilla AIOS ruta=" + plantillaAios + "; TRM ruta=" + pibPeaTrmDg + ".");
        explicaciones.put(64, "valor = patrimonioUsd / mensual.afiliados() * 1,000,000; patrimonioUsd=(activos-pasivos)/TRM desde CUENTAS ruta=" + plantillaAios + " y afiliados por query Teradata Formato491.");
        explicaciones.put(65, "valor = cuentas.resultadoNeto() / fila 51 * 100; resultado neto desde Plantilla AIOS hoja CUENTAS y comisiones desde la query 411500; plantilla=" + plantillaAios + ".");
        explicaciones.put(66, "valor = cuentas.resultadoNeto() / patrimonioUsd * 100; resultado neto desde CUENTAS ruta=" + plantillaAios + " y patrimonioUsd=(activos-pasivos)/TRM.");
        explicaciones.put(67, "valor = fila 52 / mensual.afiliados() * 1,000,000; gastos desde query Teradata de la fila 52; plantilla=" + plantillaAios + "; afiliados por query Teradata Formato491.");
        explicaciones.put(68, "valor = fila 51 / mensual.aportantesSemestral() * 1,000,000; comisiones desde la query Teradata 411500; plantilla=" + plantillaAios + "; aportantes semestrales desde query Teradata.");
        explicaciones.put(69, "valor = cuentas.admon() / fila 61; administración desde CUENTAS ruta=" + plantillaAios + " y fila 61 calculada con " + queryAportes136 + ".");
        explicaciones.put(70, "valor fijo = 16; no usa insumo externo.");
        explicaciones.put(71, "valor = promedio(trimestral.comisionesPct col_obl, por_obl, pro_obl, ska_obl); valores obtenidos por OCR de la Carta Circular SFC correspondiente al período de corte; el log muestra el PDF, el texto OCR y las cuatro comisiones utilizadas.");
        explicaciones.put(72, "valor fijo = 0; no usa insumo externo.");
        explicaciones.put(73, "valor fijo = 0; no usa insumo externo.");
        explicaciones.put(74, "valor = (3 - fila 71) * 0.25; usa comisión promedio porcentual calculada en fila 71.");
        explicaciones.put(75, "valor = (3 - fila 71) * 0.75; usa comisión promedio porcentual calculada en fila 71.");
        explicaciones.put(76, "valor fijo = 0; no usa insumo externo.");
        explicaciones.put(77, "valor = fila 51; comisiones desde la query Teradata 411500; plantilla=" + plantillaAios + ".");
        explicaciones.put(78, "valor = fila 28; se reutiliza el valor de fondos administrados consultado en Teradata y convertido con TRM ruta=" + pibPeaTrmDg + ".");
        explicaciones.put(79, "valor = fila 77 / fila 78; fila 77 son comisiones desde la query Teradata 411500; plantilla=" + plantillaAios + " y fila 78 fondos administrados.");
        explicaciones.put(80, "valor = año(fechaCorte) - 1994; no usa insumo externo.");
        explicaciones.put(82, "valor = rentabilidad nominal 10 años calculada por RentabilidadService * 100 para expresarla en puntos porcentuales sin símbolo %; usa NAV del archivo Valores_Fondo_Moder/MODERADO ruta=" + valoresFondo + " y fechas de corte.");
        explicaciones.put(83, "valor = rentabilidad real 10 años calculada por RentabilidadService * 100 para expresarla en puntos porcentuales sin símbolo %; usando rentabilidad nominal, IPC y archivo Rent_Vr_Uni_Moderado ruta=" + rentVrUni + ".");
        explicaciones.put(84, "valor = rentabilidad nominal 5 años calculada por RentabilidadService * 100 para expresarla en puntos porcentuales sin símbolo %; usando NAV de Valores_Fondo_Moder/MODERADO ruta=" + valoresFondo + ".");
        explicaciones.put(85, "valor = rentabilidad real 5 años calculada por RentabilidadService * 100 para expresarla en puntos porcentuales sin símbolo %; usando IPC/Rent_Vr_Uni_Moderado ruta=" + rentVrUni + ".");
        explicaciones.put(86, "valor = rentabilidad nominal 3 años calculada por RentabilidadService * 100 para expresarla en puntos porcentuales sin símbolo %; usando NAV de Valores_Fondo_Moder/MODERADO ruta=" + valoresFondo + ".");
        explicaciones.put(87, "valor = rentabilidad real 3 años calculada por RentabilidadService * 100 para expresarla en puntos porcentuales sin símbolo %; usando IPC/Rent_Vr_Uni_Moderado ruta=" + rentVrUni + ".");
        explicaciones.put(88, "valor = rentabilidad nominal 1 año calculada por RentabilidadService * 100 para expresarla en puntos porcentuales sin símbolo %; usando NAV de Valores_Fondo_Moder/MODERADO ruta=" + valoresFondo + ".");
        explicaciones.put(89, "valor = rentabilidad real 1 año calculada por RentabilidadService * 100 para expresarla en puntos porcentuales sin símbolo %; usando IPC/Rent_Vr_Uni_Moderado ruta=" + rentVrUni + ".");
        explicaciones.forEach((fila, explicacion) -> {
            String detalleValores = detalleValoresFila(fila, hoja, col, mensual, trimestral);
            String detalleFuente = detallesFilas.get(fila);
            String explicacionCompleta = unirPartesExplicacion(explicacion, detalleValores, detalleFuente);
            log.info("Semestral fila número {}: Explicación=\"{}\" valor={} fechaCorte={} columnaDestino={}",
                    fila, explicacionCompleta, num(hoja, fila, col), fechaCorte, col);
        });
    }

    private String unirPartesExplicacion(String... partes) {
        StringBuilder sb = new StringBuilder();
        for (String parte : partes) {
            if (parte == null || parte.isBlank()) {
                continue;
            }
            if (!sb.isEmpty()) {
                sb.append(' ');
            }
            sb.append(parte.trim());
        }
        return sb.toString();
    }

    private String detalleValoresFila(int fila, Sheet hoja, int col, MensualData mensual, TrimestralData trimestral) {
        return switch (fila) {
            case 3 -> "valores tomados: afiliados=" + mensual.afiliados() + ".";
            case 4 -> "valores tomados: afiliadosMenor30=" + mensual.afiliadosMenor30() + "; afiliados=" + mensual.afiliados() + ".";
            case 5 -> "valores tomados: afiliados30a44=" + mensual.afiliados30a44() + "; afiliados=" + mensual.afiliados() + ".";
            case 6 -> "valores tomados: afiliados45a59=" + mensual.afiliados45a59() + "; afiliados=" + mensual.afiliados() + ".";
            case 7 -> "valores tomados: afiliadosMayor60=" + mensual.afiliadosMayor60() + "; afiliados=" + mensual.afiliados() + ".";
            case 8 -> "valores tomados: constante=100.";
            case 9 -> "valores tomados: afiliados=" + mensual.afiliados() + "; divisor=1000.";
            case 10 -> "valores tomados: mujeresAfiliadasQuery=" + mensual.mujeres() + "; afiliadosTotalQuery=" + mensual.afiliados() + ".";
            case 11 -> "valores tomados: aportantesSemestral=" + mensual.aportantesSemestral() + ".";
            case 12 -> "valores tomados: afiliados=" + mensual.afiliados() + "; PEA=" + mensual.pea() + ".";
            case 13 -> "valores tomados: aportantesSemestral=" + mensual.aportantesSemestral() + "; PEA=" + mensual.pea() + ".";
            case 14 -> "valores tomados: aportantesSemestral=" + mensual.aportantesSemestral() + "; afiliados=" + mensual.afiliados() + ".";
            case 15 -> "valores tomados: salarioMinimoPonderadoCop=" + smCop(mensual) + "; salarioMinimoUsd=" + mensual.smColombiaUsd() + "; TRM=" + trm(mensual) + ".";
            case 16 -> "valores tomados: totalPensionados=query Formato495=" + num(hoja, 16, col) + ".";
            case 17 -> "valores tomados: invalidez=" + mensual.totalInv() + "; totalPensionados=fila16=" + num(hoja, 16, col) + ".";
            case 18 -> "valores tomados: vejez=" + mensual.totalVej() + "; totalPensionados=fila16=" + num(hoja, 16, col) + ".";
            case 19 -> "valores tomados: sobrevivencia=" + mensual.totalSob() + "; totalPensionados=fila16=" + num(hoja, 16, col) + ".";
            case 25 -> "valores tomados: fila25=" + num(hoja, 25, col) + "; fuente=query Teradata Formato493 fallecidos dividido entre 1000.";
            case 26 -> "valores tomados: traspasosSistema=" + mensual.traspasosSistema() + ".";
            case 27 -> "valores tomados: traspasosSistema=" + mensual.traspasosSistema() + "; afiliados=" + mensual.afiliados() + ".";
            case 28 -> "valores tomados: fondoAdministradoMmCop=" + mensual.vrFondo() + "; TRM=" + trm(mensual) + "; operación=fondoAdministradoMmCop/TRM.";
            case 29 -> "valores tomados: fila28=" + num(hoja, 28, col) + "; pibSemestralCOP=" + mensual.pibSemestral() + "; TRM=" + trm(mensual) + "; pibUsd=" + safeDivide(mensual.pibSemestral(), trm(mensual)) + ".";
            case 30 -> "valores tomados: total1=" + mensual.total1() + "; TRM=" + trm(mensual) + ".";
            case 31 -> "valores tomados: dudaG=" + mensual.dudaG() + ".";
            case 32 -> "valores tomados: dudaEf=" + mensual.dudaEf() + ".";
            case 33 -> "valores tomados: dudaNf=" + mensual.dudaNf() + ".";
            case 34 -> "valores tomados: dudaAc=" + mensual.dudaAc() + ".";
            case 35 -> "valores tomados: dudaF=" + mensual.dudaF() + ".";
            case 36 -> "valores tomados: constante=0.";
            case 37 -> "valores tomados: dudaGe=" + mensual.dudaGe() + ".";
            case 38 -> "valores tomados: dudaEfe=" + mensual.dudaEfe() + ".";
            case 39 -> "valores tomados: dudaNfe=" + mensual.dudaNfe() + ".";
            case 40 -> "valores tomados: dudaAce=" + mensual.dudaAce() + ".";
            case 41 -> "valores tomados: dudaFe=" + mensual.dudaFe() + ".";
            case 42 -> "valores tomados: constante=2.";
            case 43 -> "valores tomados: otros=" + mensual.otros() + ".";
            case 44 -> "valores tomados: fila44=" + num(hoja, 44, col) + ".";
            case 45 -> "valores tomados: fila28=" + num(hoja, 28, col) + "; fila45=" + num(hoja, 45, col) + ".";
            case 46 -> "valores tomados: constante=4.";
            case 47 -> "valores tomados: fila47=" + num(hoja, 47, col) + "; fallback mensual.porcVrFondo=" + mensual.porcVrFondo() + ".";
            case 48 -> "valores tomados: activosCuentas=" + mensual.activosCuentas() + "; TRM=" + trm(mensual) + ".";
            case 49 -> "valores tomados: pasivosCuentas=" + mensual.pasivosCuentas() + "; TRM=" + trm(mensual) + ".";
            case 50 -> "valores tomados: activosCuentas=" + mensual.activosCuentas() + "; pasivosCuentas=" + mensual.pasivosCuentas() + "; TRM=" + trm(mensual) + ".";
            case 51 -> "valores tomados: comisiones=fila51=" + num(hoja, 51, col) + ".";
            case 52 -> "valores tomados: gastos=fila52=" + num(hoja, 52, col) + ".";
            case 53 -> "valores tomados: resultadoOperacion=fila53=" + num(hoja, 53, col) + ".";
            case 54 -> "valores tomados: resultadoNeto=fila54=" + num(hoja, 54, col) + ".";
            case 55 -> "valores tomados: administracion=fila55=" + num(hoja, 55, col) + ".";
            case 56 -> "valores tomados: cuenta511500/TRM=fila56=" + num(hoja, 56, col) + "; TRM=" + trm(mensual) + ".";
            case 57 -> "valores tomados: cuenta511527/TRM=fila57=" + num(hoja, 57, col) + "; TRM=" + trm(mensual) + ".";
            case 58 -> "valores tomados: fila56=" + num(hoja, 56, col) + "; fila57=" + num(hoja, 57, col) + "; TRM=" + trm(mensual) + ".";
            case 59 -> "valores tomados: suma otros gastos/TRM=fila59=" + num(hoja, 59, col) + "; TRM=" + trm(mensual) + ".";
            case 60 -> "valores tomados: cuenta510000/TRM=fila60=" + num(hoja, 60, col) + "; TRM=" + trm(mensual) + ".";
            case 61 -> "valores tomados: aportantesSemestral=" + mensual.aportantesSemestral() + "; TRM=" + trm(mensual) + "; fila61=" + num(hoja, 61, col) + ".";
            case 62 -> "valores tomados: gastos=fila52=" + num(hoja, 52, col) + "; fila62=" + num(hoja, 62, col) + ".";
            case 63 -> "valores tomados: fila28=" + num(hoja, 28, col) + "; fila63=" + num(hoja, 63, col) + "; TRM=" + trm(mensual) + ".";
            case 64 -> "valores tomados: patrimonioUsd=fila50=" + num(hoja, 50, col) + "; afiliados=" + mensual.afiliados() + ".";
            case 65 -> "valores tomados: resultadoNeto=fila54=" + num(hoja, 54, col) + "; comisiones=fila51=" + num(hoja, 51, col) + ".";
            case 66 -> "valores tomados: resultadoNeto=fila54=" + num(hoja, 54, col) + "; patrimonioUsd=fila50=" + num(hoja, 50, col) + ".";
            case 67 -> "valores tomados: gastos=fila52=" + num(hoja, 52, col) + "; afiliados=" + mensual.afiliados() + ".";
            case 68 -> "valores tomados: comisiones=fila51=" + num(hoja, 51, col) + "; aportantesSemestral=" + mensual.aportantesSemestral() + ".";
            case 69 -> "valores tomados: administracion=fila55=" + num(hoja, 55, col) + "; fila61=" + num(hoja, 61, col) + ".";
            case 70 -> "valores tomados: constante=16.";
            case 71 -> "valores tomados: col_obl=" + trimestral.comisionesPct().getOrDefault("col_obl", BigDecimal.ZERO) + "; por_obl=" + trimestral.comisionesPct().getOrDefault("por_obl", BigDecimal.ZERO) + "; pro_obl=" + trimestral.comisionesPct().getOrDefault("pro_obl", BigDecimal.ZERO) + "; ska_obl=" + trimestral.comisionesPct().getOrDefault("ska_obl", BigDecimal.ZERO) + ".";
            case 72 -> "valores tomados: constante=0.";
            case 73 -> "valores tomados: constante=0.";
            case 74 -> "valores tomados: constante=3; fila71=" + num(hoja, 71, col) + "; factor=0.25.";
            case 75 -> "valores tomados: constante=3; fila71=" + num(hoja, 71, col) + "; factor=0.75.";
            case 76 -> "valores tomados: constante=0.";
            case 77 -> "valores tomados: comisiones=fila77=" + num(hoja, 77, col) + ".";
            case 78 -> "valores tomados: fila28=" + num(hoja, 28, col) + ".";
            case 79 -> "valores tomados: fila77=" + num(hoja, 77, col) + "; fila78=" + num(hoja, 78, col) + ".";
            case 80 -> "valores tomados: anioFechaCorte=" + fechaYearFromColumnValue(hoja, col) + "; base=1994.";
            case 82 -> "valores tomados: rentabilidadNominal10=" + num(hoja, 82, col) + ".";
            case 83 -> "valores tomados: rentabilidadReal10=" + num(hoja, 83, col) + ".";
            case 84 -> "valores tomados: rentabilidadNominal5=" + num(hoja, 84, col) + ".";
            case 85 -> "valores tomados: rentabilidadReal5=" + num(hoja, 85, col) + ".";
            case 86 -> "valores tomados: rentabilidadNominal3=" + num(hoja, 86, col) + ".";
            case 87 -> "valores tomados: rentabilidadReal3=" + num(hoja, 87, col) + ".";
            case 88 -> "valores tomados: rentabilidadNominal1=" + num(hoja, 88, col) + ".";
            case 89 -> "valores tomados: rentabilidadReal1=" + num(hoja, 89, col) + ".";
            default -> "";
        };
    }

    private int fechaYearFromColumnValue(Sheet hoja, int col) {
        BigDecimal fila80 = num(hoja, 80, col);
        return fila80.add(BigDecimal.valueOf(1994)).intValue();
    }

    private String rutaInsumo(String nombre, PathSupplier supplier) {
        try {
            Path path = supplier.get();
            return path == null ? nombre + " (ruta no resuelta)" : path.toAbsolutePath().toString();
        } catch (Exception e) {
            return nombre + " (ruta no resuelta: " + e.getMessage() + ")";
        }
    }

    @FunctionalInterface
    private interface PathSupplier {
        Path get() throws Exception;
    }

    private BigDecimal promedioComisionObligatoria(TrimestralData trimestral) {
        BigDecimal col = trimestral.comisionesPct().getOrDefault("col_obl", BigDecimal.ZERO);
        BigDecimal por = trimestral.comisionesPct().getOrDefault("por_obl", BigDecimal.ZERO);
        BigDecimal pro = trimestral.comisionesPct().getOrDefault("pro_obl", BigDecimal.ZERO);
        BigDecimal ska = trimestral.comisionesPct().getOrDefault("ska_obl", BigDecimal.ZERO);
        return col.add(por).add(pro).add(ska).divide(BigDecimal.valueOf(4), 8, RoundingMode.HALF_UP);
    }


    private BigDecimal readTotalPensionados495(LocalDate fechaCorte, BigDecimal fallback) {
        try {
            Path file495 = findPensionados495File(fechaCorte);
            try (Workbook wb = WorkbookFactory.create(file495.toFile(), null, true)) {
                Sheet totalPensionados = getSheetIgnoreCase(wb, "TOTAL PENSIONADOS");
                if (totalPensionados == null) {
                    totalPensionados = findSheetContainsIgnoreCase(wb, "total pensionados");
                }
                if (totalPensionados == null) {
                    log.warn("Semestral fila16: no se encontró hoja TOTAL PENSIONADOS en {}. Se usa fallback={}",
                            file495.toAbsolutePath(), fallback);
                    return fallback == null ? BigDecimal.ZERO : fallback;
                }

                setDate(totalPensionados, "B4", fechaCorte);
                BigDecimal valor = readTotalPensionadosSerie(totalPensionados, fechaCorte);
                if (valor.signum() != 0) {
                    log.info("Semestral fila16: archivo={} hoja=TOTAL PENSIONADOS parámetro=B4 fecha={} valor columna I={}",
                            file495.toAbsolutePath(), fechaCorte, valor);
                    return valor;
                }
                log.warn("Semestral fila16: no se encontró valor en columna I para fecha={} en {}. Se usa fallback={}",
                        fechaCorte, file495.toAbsolutePath(), fallback);
            }
        } catch (Exception e) {
            log.warn("Semestral fila16: no fue posible leer Series_Formato-495 PENSIONADOS para fecha={}: {}. Se usa fallback={}",
                    fechaCorte, e.getMessage(), fallback);
        }
        return fallback == null ? BigDecimal.ZERO : fallback;
    }


    private PensionadosPorEntidad readPensionadosPorEntidad495(LocalDate fechaCorte, PensionadosPorEntidad fallback) {
        try {
            Path file495 = findPensionados495File(fechaCorte);
            try (Workbook wb = WorkbookFactory.create(file495.toFile(), null, true)) {
                FormulaEvaluator evaluator = wb.getCreationHelper().createFormulaEvaluator();
                Sheet porEntidad = getSheetIgnoreCase(wb, "por Entidad");
                if (porEntidad == null) {
                    porEntidad = findSheetContainsIgnoreCase(wb, "por entidad");
                }
                if (porEntidad == null) {
                    log.warn("Semestral filas17-19: no se encontró hoja por Entidad en {}. Se usa fallback={}",
                            file495.toAbsolutePath(), fallback);
                    return fallback;
                }

                setDate(porEntidad, "C6", fechaCorte);
                evaluator.clearAllCachedResultValues();
                BigDecimal invalidez = num(porEntidad, "BI62", evaluator);
                BigDecimal vejez = num(porEntidad, "BH62", evaluator);
                BigDecimal sobrevivencia = num(porEntidad, "BJ62", evaluator);
                log.info("Semestral filas17-19: archivo={} hoja=por Entidad parámetro=C6 fecha={} BI62(invalidez)={} BH62(vejez)={} BJ62(sobrevivencia)={}",
                        file495.toAbsolutePath(), fechaCorte, invalidez, vejez, sobrevivencia);
                return new PensionadosPorEntidad(
                        invalidez.signum() == 0 ? fallback.invalidez() : invalidez,
                        vejez.signum() == 0 ? fallback.vejez() : vejez,
                        sobrevivencia.signum() == 0 ? fallback.sobrevivencia() : sobrevivencia
                );
            }
        } catch (Exception e) {
            log.warn("Semestral filas17-19: no fue posible leer por Entidad BI62/BH62/BJ62 para fecha={}: {}. Se usa fallback={}",
                    fechaCorte, e.getMessage(), fallback);
            return fallback;
        }
    }

    private Path findPensionados495File(LocalDate fechaCorte) {
        Path principal = properties.insumosDir()
                .resolve("Formato 495")
                .resolve("Series_Formato-495 PENSIONADOS.xlsm");
        if (Files.isRegularFile(principal)) {
            return principal;
        }
        try {
            return locator.findRequired("Series_Formato-495 PENSIONADOS", fechaCorte);
        } catch (Exception ignored) {
            Path local = Path.of("insumos_ejemplo", "Series_Formato-495 PENSIONADOS.xlsm");
            if (Files.isRegularFile(local)) return local;
            throw ignored;
        }
    }

    private BigDecimal readTotalPensionadosSerie(Sheet totalPensionados, LocalDate fechaCorte) {
        DataFormatter formatter = new DataFormatter(Locale.forLanguageTag("es-CO"));
        BigDecimal mejor = BigDecimal.ZERO;
        LocalDate mejorFecha = LocalDate.MIN;
        int mejorFila = -1;
        for (int r = 0; r <= totalPensionados.getLastRowNum(); r++) {
            Row row = totalPensionados.getRow(r);
            if (row == null) continue;
            LocalDate fechaFila = cellAsDate(row.getCell(1)); // columna B
            if (fechaFila == null) continue;
            BigDecimal valor = cellAsBigDecimal(row.getCell(8), formatter); // columna I
            if (valor.signum() == 0) continue;
            if (fechaFila.equals(fechaCorte)) {
                log.info("Semestral fila16: match exacto hoja=TOTAL PENSIONADOS fila={} fecha={} celda=I{} valor={}",
                        r + 1, fechaFila, r + 1, valor);
                return valor;
            }
            if (!fechaFila.isAfter(fechaCorte) && fechaFila.isAfter(mejorFecha)) {
                mejorFecha = fechaFila;
                mejor = valor;
                mejorFila = r + 1;
            }
        }
        if (mejor.signum() != 0) {
            log.info("Semestral fila16: usando fecha anterior hoja=TOTAL PENSIONADOS fila={} fechaFila={} fechaCorte={} celda=I{} valor={}",
                    mejorFila, mejorFecha, fechaCorte, mejorFila, mejor);
        }
        return mejor;
    }

    private BigDecimal cellAsBigDecimal(Cell cell, DataFormatter formatter) {
        if (cell == null) return BigDecimal.ZERO;
        try {
            if (cell.getCellType() == CellType.NUMERIC ||
                    (cell.getCellType() == CellType.FORMULA && cell.getCachedFormulaResultType() == CellType.NUMERIC)) {
                return BigDecimal.valueOf(cell.getNumericCellValue());
            }
            String text = formatter.formatCellValue(cell);
            if (text == null || text.isBlank()) return BigDecimal.ZERO;
            String normalized = text.trim().replace("$", "").replace("%", "").replace(" ", "");
            if (normalized.contains(",") && normalized.contains(".")) {
                normalized = normalized.replace(",", "");
            } else if (normalized.contains(",")) {
                normalized = normalized.replace(".", "").replace(",", ".");
            }
            return new BigDecimal(normalized);
        } catch (Exception e) {
            return BigDecimal.ZERO;
        }
    }

    private void setDate(Sheet sheet, String ref, LocalDate date) {
        Cell target = cell(sheet, ref);
        target.setCellValue(Date.from(date.atStartOfDay(ZoneId.systemDefault()).toInstant()));
    }

    private Sheet findSheetContainsIgnoreCase(Workbook wb, String fragment) {
        String normalizedFragment = fragment.toLowerCase(Locale.ROOT);
        for (int i = 0; i < wb.getNumberOfSheets(); i++) {
            Sheet sheet = wb.getSheetAt(i);
            if (sheet.getSheetName().toLowerCase(Locale.ROOT).contains(normalizedFragment)) {
                return sheet;
            }
        }
        return null;
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

    int columnaSemestral(Sheet hoja, LocalDate fechaCorte) {
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

        int lastPeriodCol = 1;
        for (int c = 2; c < Math.max(last, 3); c++) {
            String mes = fmt.formatCellValue(rowMes.getCell(c));
            String anio = fmt.formatCellValue(rowAnio.getCell(c));
            if ((mes != null && !mes.isBlank()) || (anio != null && !anio.isBlank())) lastPeriodCol = c;
        }
        int targetColIndex = Math.max(lastPeriodCol + 1, 2);
        copyPreviousColumnFormat(hoja, targetColIndex);
        Row targetMes = hoja.getRow(0);
        Row targetAnio = hoja.getRow(1);
        Cell mesCell = targetMes.getCell(targetColIndex);
        if (mesCell == null) mesCell = targetMes.createCell(targetColIndex);
        Cell anioCell = targetAnio.getCell(targetColIndex);
        if (anioCell == null) anioCell = targetAnio.createCell(targetColIndex);
        mesCell.setCellValue(mesObjetivo);
        anioCell.setCellValue(fechaCorte.getYear());
        return targetColIndex + 1;
    }

    private void copyPreviousColumnFormat(Sheet sheet, int targetColIndex) {
        int sourceColIndex = targetColIndex - 1;
        if (sourceColIndex < 0) return;
        sheet.setColumnWidth(targetColIndex, sheet.getColumnWidth(sourceColIndex));
        sheet.setColumnHidden(targetColIndex, sheet.isColumnHidden(sourceColIndex));

        for (int rowIndex = 0; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
            Row row = sheet.getRow(rowIndex);
            if (row == null) continue;
            Cell sourceCell = row.getCell(sourceColIndex);
            if (sourceCell == null) continue;
            Cell targetCell = row.getCell(targetColIndex);
            if (targetCell == null) targetCell = row.createCell(targetColIndex);
            targetCell.setCellStyle(sourceCell.getCellStyle());
        }
    }

    void normalizarEstilosSemestral(Sheet sheet) {
        Workbook workbook = sheet.getWorkbook();
        Map<Integer, Font> blackFonts = new HashMap<>();
        Map<String, CellStyle> normalizedStyles = new HashMap<>();
        for (Row row : sheet) {
            for (Cell cell : row) {
                CellStyle original = cell.getCellStyle();
                boolean removeBandFill = row.getRowNum() == 80 && cell.getColumnIndex() >= 2;
                String key = original.getIndex() + ":" + removeBandFill;
                CellStyle normalized = normalizedStyles.computeIfAbsent(key, ignored -> {
                    CellStyle style = workbook.createCellStyle();
                    style.cloneStyleFrom(original);
                    Font font = blackFonts.computeIfAbsent((int) original.getFontIndex(), fontIndex ->
                            cloneFontInBlack(workbook, workbook.getFontAt(fontIndex)));
                    style.setFont(font);
                    if (removeBandFill) {
                        style.setFillPattern(FillPatternType.NO_FILL);
                        style.setFillForegroundColor((short) 0);
                        style.setFillBackgroundColor((short) 0);
                    }
                    return style;
                });
                cell.setCellStyle(normalized);
            }
        }
    }

    private Font cloneFontInBlack(Workbook workbook, Font source) {
        Font target = workbook.createFont();
        target.setFontName(source.getFontName());
        target.setFontHeight(source.getFontHeight());
        target.setBold(source.getBold());
        target.setItalic(source.getItalic());
        target.setStrikeout(source.getStrikeout());
        target.setUnderline(source.getUnderline());
        target.setTypeOffset(source.getTypeOffset());
        target.setCharSet(source.getCharSet());
        target.setColor(IndexedColors.BLACK.getIndex());
        return target;
    }

    private String normalize(String value) {
        return value == null ? "" : value.trim().toLowerCase(Locale.ROOT);
    }

    private BigDecimal trm(MensualData data) {
        return data.trm().signum() == 0 ? BigDecimal.ONE : data.trm();
    }

    private BigDecimal smCop(MensualData data) {
        return data.smColombiaUsd().multiply(trm(data));
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
                    num(cuentas, "C22"),
                    num(cuentas, "C24")
                            .add(num(cuentas, "C28"))
                            .add(num(cuentas, "C29"))
                            .add(num(cuentas, "C31"))
                            .add(num(cuentas, "C32"))
                            .add(num(cuentas, "C33"))
                            .add(num(cuentas, "C34"))
                            .add(num(cuentas, "C35"))
                            .add(num(cuentas, "C36"))
                            .add(num(cuentas, "C37"))
                            .add(num(cuentas, "C38")),
                    num(cuentas, "C15")
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
        BigDecimal value = formato136QueryService.leerAportesRecibidos(fechaCorte);
        LocalDate fechaInicial = fechaCorte.minusYears(1).withDayOfMonth(1);
        log.info("Semestral: aportes recibidos Formato 136 desde query Teradata={} para fecha={} ventana={}..{}",
                value, fechaCorte, fechaInicial, fechaCorte);
        return value;
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

            Set<String> keysEsperadasBaseMes = new HashSet<>();
            for (String entidad : entidades) {
                keysEsperadasBaseMes.add(entidad + "-" + serialFecha + "-" + cuentaPatrimonio);
            }
            BigDecimal sumaCop = BigDecimal.ZERO;
            Set<String> encontradas = new HashSet<>();
            int last = baseMes.getLastRowNum() + 1;
            for (int r = 2; r <= last; r++) {
                Row row = baseMes.getRow(r - 1);
                if (row == null) continue;
                String entidad = normalize(cellAsString(row.getCell(4))); // col E
                if (!entidades.contains(entidad)) continue;
                String cuenta = normalize(cellAsString(row.getCell(1))); // col B
                if (!cuentaPatrimonio.equals(cuenta)) continue;
                Integer serialFila = excelSerialFromCell(row.getCell(3)); // col D
                if (serialFila == null || serialFila != serialFecha) continue;
                String keyConstruida = entidad + "-" + serialFila + "-" + cuenta;
                if (!keysEsperadasBaseMes.contains(keyConstruida)) continue;
                BigDecimal valor = num(baseMes, r, 6); // col F valor
                if (valor.signum() > 0) {
                    sumaCop = sumaCop.add(valor);
                    encontradas.add(keyConstruida);
                    log.info("Patrimonio base mes match: key={} valorCOP={}", keyConstruida, valor);
                }
                if (encontradas.size() == keysEsperadasBaseMes.size()) break;
            }
            BigDecimal mmCop = sumaCop.divide(BigDecimal.valueOf(1_000_000), 8, RoundingMode.HALF_UP);
            log.info("Patrimonio base mes total: fechaParametro={} fechaBaseMes={} serialBaseMes={} serialFechaCorte={} entidades={} matches={} sumaCOP={} sumaMMCOP={}",
                    fechaCorte, fechaBaseMes, serialFecha, serialFechaCorte, entidades, encontradas.size(), sumaCop, mmCop);
            if (encontradas.size() < keysEsperadasBaseMes.size()) {
                log.warn("Patrimonio base mes incompleto: esperadas={} encontradas={} faltantes={}",
                        keysEsperadasBaseMes.size(), encontradas.size(), keysEsperadasBaseMes.stream().filter(k -> !encontradas.contains(k)).toList());
                // Fallback defensivo: si no hay match con serial del primer día de mes, intentar serial exacto de fecha de corte.
                if (serialFechaCorte != serialFecha) {
                    Set<String> keysEsperadasCorte = new HashSet<>();
                    for (String entidad : entidades) {
                        keysEsperadasCorte.add(entidad + "-" + serialFechaCorte + "-" + cuentaPatrimonio);
                    }
                    BigDecimal sumaCopCorte = BigDecimal.ZERO;
                    Set<String> encontradasCorte = new HashSet<>();
                    for (int r = 2; r <= last; r++) {
                        Row row = baseMes.getRow(r - 1);
                        if (row == null) continue;
                        String entidad = normalize(cellAsString(row.getCell(4)));
                        if (!entidades.contains(entidad)) continue;
                        String cuenta = normalize(cellAsString(row.getCell(1)));
                        if (!cuentaPatrimonio.equals(cuenta)) continue;
                        Integer serialFila = excelSerialFromCell(row.getCell(3));
                        if (serialFila == null || serialFila != serialFechaCorte) continue;
                        String keyConstruida = entidad + "-" + serialFila + "-" + cuenta;
                        if (!keysEsperadasCorte.contains(keyConstruida)) continue;
                        BigDecimal valor = num(baseMes, r, 6);
                        if (valor.signum() > 0) {
                            sumaCopCorte = sumaCopCorte.add(valor);
                            encontradasCorte.add(keyConstruida);
                            log.info("Patrimonio base mes fallback(fecha corte) match: key={} valorCOP={}", keyConstruida, valor);
                        }
                        if (encontradasCorte.size() == keysEsperadasCorte.size()) break;
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

    private String cellAsString(Cell cell) {
        if (cell == null) return "";
        try {
            return switch (cell.getCellType()) {
                case STRING -> cell.getStringCellValue();
                case NUMERIC -> BigDecimal.valueOf(cell.getNumericCellValue()).stripTrailingZeros().toPlainString();
                case FORMULA -> cell.getCachedFormulaResultType() == org.apache.poi.ss.usermodel.CellType.NUMERIC
                        ? BigDecimal.valueOf(cell.getNumericCellValue()).stripTrailingZeros().toPlainString()
                        : cell.getRichStringCellValue().getString();
                default -> cell.toString();
            };
        } catch (Exception e) {
            return cell.toString();
        }
    }

    private Integer excelSerialFromCell(Cell cell) {
        if (cell == null) return null;
        try {
            return switch (cell.getCellType()) {
                case NUMERIC -> (int) Math.round(cell.getNumericCellValue());
                case FORMULA -> cell.getCachedFormulaResultType() == org.apache.poi.ss.usermodel.CellType.NUMERIC
                        ? (int) Math.round(cell.getNumericCellValue())
                        : null;
                default -> null;
            };
        } catch (Exception e) {
            return null;
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



    private DatoDetalle readDeudaGubernamentalTotal(LocalDate fechaCorte) {
        try {
            Path seriesFile = locator.findRequired("PIB_PEA_TRM_DG", fechaCorte);
            try (Workbook wb = WorkbookFactory.create(seriesFile.toFile(), null, true)) {
                Sheet sheet = wb.getSheet("Hoja1");
                if (sheet == null) sheet = wb.getSheetAt(0);
                BigDecimal mejor = BigDecimal.ZERO;
                LocalDate mejorFecha = LocalDate.MIN;
                int mejorFila = -1;
                for (Row row : sheet) {
                    LocalDate fecha = cellAsDate(row.getCell(11)); // columna L
                    if (fecha == null || fecha.isAfter(fechaCorte)) continue;
                    BigDecimal deuda = num(sheet, row.getRowNum() + 1, 13); // columna M
                    if (deuda.signum() == 0) continue;
                    if (fecha.equals(fechaCorte)) {
                        String detalle = "detalle fuente deuda gubernamental: archivo=" + seriesFile.toAbsolutePath()
                                + " hoja=Hoja1 fila=" + (row.getRowNum() + 1) + " fecha=" + fecha
                                + " celda=M" + (row.getRowNum() + 1) + " valor=" + deuda + ".";
                        log.debug("Deuda gubernamental total exacta: {}", detalle);
                        return new DatoDetalle(deuda, detalle);
                    }
                    if (fecha.isAfter(mejorFecha)) {
                        mejorFecha = fecha;
                        mejor = deuda;
                        mejorFila = row.getRowNum() + 1;
                    }
                }
                if (mejor.signum() != 0) {
                    String detalle = "detalle fuente deuda gubernamental: archivo=" + seriesFile.toAbsolutePath()
                            + " hoja=Hoja1 fila=" + mejorFila + " fechaFila=" + mejorFecha + " fechaCorte=" + fechaCorte
                            + " celda=M" + mejorFila + " valor=" + mejor + " (fallback por fecha anterior).";
                    log.debug("Deuda gubernamental total por fecha anterior: {}", detalle);
                    return new DatoDetalle(mejor, detalle);
                } else {
                    String detalle = "no se encontró deuda gubernamental total en archivo=" + seriesFile.toAbsolutePath()
                            + " hoja=Hoja1 columna M para fecha=" + fechaCorte + " o anterior.";
                    log.warn("Semestral fila45: {}", detalle);
                    return new DatoDetalle(BigDecimal.ZERO, detalle);
                }
            }
        } catch (Exception e) {
            String detalle = "no fue posible leer deuda gubernamental total desde PIB_PEA_TRM_DG para fecha=" + fechaCorte + ": " + e.getMessage() + ".";
            log.warn("Semestral fila45: {}", detalle);
            return new DatoDetalle(BigDecimal.ZERO, detalle);
        }
    }

    private DatoDetalle readFila44DesdeLimites(LocalDate fechaCorte) {
        try {
            Path limites = locator.findRequired("LIMITES", fechaCorte);
            try (Workbook wb = WorkbookFactory.create(limites.toFile(), null, true)) {
                Sheet aios = getSheetIgnoreCase(wb, "AIOS");
                if (aios == null) {
                    String detalle = "archivo=" + limites.toAbsolutePath() + " hoja=AIOS no encontrada.";
                    return new DatoDetalle(BigDecimal.ZERO, detalle);
                }
                BigDecimal o4 = num(aios, "O4", null);
                BigDecimal q4 = num(aios, "Q4", null);
                BigDecimal s4 = num(aios, "S4", null);
                BigDecimal u4 = num(aios, "U4", null);
                BigDecimal w4 = num(aios, "W4", null);
                BigDecimal y4 = num(aios, "Y4", null);
                BigDecimal suma = o4.add(q4).add(s4).add(u4).add(w4).add(y4);
                BigDecimal porcentaje = suma.multiply(BigDecimal.valueOf(100));
                String detalle = "detalle fuente: archivo=" + limites.toAbsolutePath()
                        + " hoja=AIOS celdas O4=" + o4 + ", Q4=" + q4 + ", S4=" + s4 + ", U4=" + u4
                        + ", W4=" + w4 + ", Y4=" + y4 + "; suma=" + suma + "; operación=suma * 100.";
                log.debug("Semestral fila44 desde LIMITES: {} resultado={}", detalle, porcentaje);
                return new DatoDetalle(porcentaje, detalle);
            }
        } catch (Exception e) {
            String detalle = "no fue posible leer LIMITES para fecha=" + fechaCorte + ": " + e.getMessage() + ".";
            log.warn("Semestral fila44: {}", detalle);
            return new DatoDetalle(BigDecimal.ZERO, detalle);
        }
    }

    private BigDecimal readFila25Trimestral493(LocalDate fechaCorte) {
        return divide(formato493QueryService.leerFallecidosSistema(fechaCorte), BigDecimal.valueOf(1000));
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

    private record DatoDetalle(BigDecimal valor, String detalle) {}

    private record PensionadosPorEntidad(BigDecimal invalidez, BigDecimal vejez, BigDecimal sobrevivencia) {}

    private record CuentasData(
            BigDecimal comisiones,
            BigDecimal gastos,
            BigDecimal resultadoOperacion,
            BigDecimal resultadoNeto,
            BigDecimal admon,
            BigDecimal cuenta511500,
            BigDecimal cuenta511527,
            BigDecimal otrosGastosOperacion,
            BigDecimal gastoOperacion510000
    ) {
        static final CuentasData ZERO = new CuentasData(
                BigDecimal.ZERO, BigDecimal.ZERO, BigDecimal.ZERO, BigDecimal.ZERO,
                BigDecimal.ZERO, BigDecimal.ZERO, BigDecimal.ZERO, BigDecimal.ZERO, BigDecimal.ZERO
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

    void writeFilasAfiliadosDisponibilidad(Sheet sheet, int column, MensualData mensual) {
        write(sheet, 3, column, mensual.afiliados());
        write(sheet, 20, column, "No Disponible");
    }

    private void write(Sheet sheet, int row1Based, int col1Based, BigDecimal value) {
        Row row = sheet.getRow(row1Based - 1);
        if (row == null) row = sheet.createRow(row1Based - 1);
        Cell cell = row.getCell(col1Based - 1);
        if (cell == null) cell = row.createCell(col1Based - 1);
        cell.setCellValue(value == null ? 0d : value.doubleValue());
    }

    private void write(Sheet sheet, int row1Based, int col1Based, String value) {
        Row row = sheet.getRow(row1Based - 1);
        if (row == null) row = sheet.createRow(row1Based - 1);
        Cell cell = row.getCell(col1Based - 1);
        if (cell == null) cell = row.createCell(col1Based - 1);
        cell.setCellValue(value);
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
