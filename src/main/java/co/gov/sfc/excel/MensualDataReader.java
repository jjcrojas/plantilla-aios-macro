package co.gov.sfc.excel;

import co.gov.sfc.config.AiosProperties;
import co.gov.sfc.insumos.InsumosLocator;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.util.IOUtils;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Component;

import java.math.BigDecimal;
import java.math.RoundingMode;
import java.time.LocalDate;
import java.time.format.TextStyle;
import java.util.Locale;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.stream.XMLInputFactory;
import javax.xml.stream.XMLStreamConstants;
import javax.xml.stream.XMLStreamReader;
import java.io.InputStream;
import java.nio.file.Path;
import java.nio.file.Files;
import java.util.zip.ZipFile;

@Component
public class MensualDataReader {

    private static final Logger log = LoggerFactory.getLogger(MensualDataReader.class);
    private final InsumosLocator locator;
    private final AiosProperties properties;
    private final Formato491QueryService formato491QueryService;
    private final FondoAdministradoQueryService fondoAdministradoQueryService;
    private final Formato493QueryService formato493QueryService;
    private final Formato495QueryService formato495QueryService;
    private final TrmService trmService;
    private final SeriesEconomicasService seriesEconomicasService;
    private final BalanceContableQueryService balanceContableQueryService;
    private final RentabilidadService rentabilidadService;

    public MensualDataReader(InsumosLocator locator, AiosProperties properties,
                             Formato491QueryService formato491QueryService,
                             FondoAdministradoQueryService fondoAdministradoQueryService,
                             Formato493QueryService formato493QueryService,
                             Formato495QueryService formato495QueryService,
                             TrmService trmService,
                             SeriesEconomicasService seriesEconomicasService,
                             BalanceContableQueryService balanceContableQueryService,
                             RentabilidadService rentabilidadService) {
        this.locator = locator;
        this.properties = properties;
        this.formato491QueryService = formato491QueryService;
        this.fondoAdministradoQueryService = fondoAdministradoQueryService;
        this.formato493QueryService = formato493QueryService;
        this.formato495QueryService = formato495QueryService;
        this.trmService = trmService;
        this.seriesEconomicasService = seriesEconomicasService;
        this.balanceContableQueryService = balanceContableQueryService;
        this.rentabilidadService = rentabilidadService;
        // Evitar asignaciones gigantes en POI que pueden terminar en OOM con archivos grandes.
        // 100 MB es suficiente para los insumos actuales y más conservador en memoria.
        IOUtils.setByteArrayMaxOverride(100_000_000);
    }

    public MensualData read(LocalDate fechaCorte) {
        log.info("Iniciando lectura de insumos para fechaCorte={}", fechaCorte);

        BigDecimal hombres = BigDecimal.ZERO;
        BigDecimal mujeres = BigDecimal.ZERO;
        BigDecimal afiliadosMenor30 = BigDecimal.ZERO;
        BigDecimal afiliados30a44 = BigDecimal.ZERO;
        BigDecimal afiliados45a59 = BigDecimal.ZERO;
        BigDecimal afiliadosMayor60 = BigDecimal.ZERO;
        BigDecimal aportantes = BigDecimal.ZERO;
        BigDecimal aportantesSemestral = BigDecimal.ZERO;
        BigDecimal afiliadosActivos = BigDecimal.ZERO;
        BigDecimal consFdosAdmon = BigDecimal.ZERO;
        BigDecimal smColombiaCop = BigDecimal.ZERO;
        BigDecimal totalPen = BigDecimal.ZERO;
        BigDecimal totalInv = BigDecimal.ZERO;
        BigDecimal totalVej = BigDecimal.ZERO;
        BigDecimal totalSob = BigDecimal.ZERO;

        // El resumen 491 se consulta una sola vez; de aquí salen todos los campos migrados a Teradata.
        final var resumen491 = formato491QueryService.leerResumen(fechaCorte);
        BigDecimal afiliadosQuery = resumen491.afiliados();
        afiliadosActivos = resumen491.afiliadosActivos();
        mujeres = resumen491.mujeresAfiliadas();
        aportantes = resumen491.aportantes();
        aportantesSemestral = resumen491.aportantesSemestral();
        consFdosAdmon = resumen491.concentracionAfiliados();
        afiliadosMenor30 = resumen491.afiliadosMenor30();
        afiliados30a44 = resumen491.afiliados30a44();
        afiliados45a59 = resumen491.afiliados45a59();
        afiliadosMayor60 = resumen491.afiliadosMayor60();
        smColombiaCop = resumen491.salarioMinimoPonderadoCop();

        BigDecimal traspasosSistema = formato493QueryService.leerTraspasosSistema(fechaCorte);

        log.info("Consulta Teradata Formato 491 completada para fechaCorte={}", fechaCorte);
        log.info("Consulta Teradata Formato 493 completada para fechaCorte={}", fechaCorte);

        var rentFile = locator.findRequired("Rent_Vr_Uni_Moderado", fechaCorte);
        var rentabilidadUnAnio = rentabilidadService.calcularRentabilidad(rentFile, fechaCorte, 1);
        BigDecimal tmpNominal1 = rentabilidadUnAnio.rentabilidadNominal();
        BigDecimal tmpReal1 = rentabilidadUnAnio.rentabilidadReal();
        log.info("Rentabilidad mensual 1 año calculada en Java con Consolidado!E e IPC_D!B: "
                        + "fechaInicio={} fechaFin={} nominal={} real={} archivo={}",
                rentabilidadUnAnio.fechaInicio(), rentabilidadUnAnio.fechaFin(),
                tmpNominal1, tmpReal1, rentFile.toAbsolutePath());

        var fondoAdministrado = fondoAdministradoQueryService.leer(fechaCorte);
        BigDecimal vrFondo = fondoAdministrado.totalMmCop();
        BigDecimal fondoSistemaJ14 = vrFondo;
        BigDecimal porcVrFondo = fondoAdministrado.concentracionProteccionPorvenirPct();
        log.info("Consulta Teradata de fondo administrado completada para fechaCorte={} totalMmCop={}",
                fechaCorte, vrFondo);

        BigDecimal total1 = BigDecimal.ZERO;
        BigDecimal dudaG = BigDecimal.ZERO;
        BigDecimal deudaGobB4 = BigDecimal.ZERO;
        BigDecimal dudaEf = BigDecimal.ZERO;
        BigDecimal dudaNf = BigDecimal.ZERO;
        BigDecimal dudaAc = BigDecimal.ZERO;
        BigDecimal dudaF = BigDecimal.ZERO;
        BigDecimal dudaGe = BigDecimal.ZERO;
        BigDecimal dudaEfe = BigDecimal.ZERO;
        BigDecimal dudaNfe = BigDecimal.ZERO;
        BigDecimal dudaAce = BigDecimal.ZERO;
        BigDecimal dudaFe = BigDecimal.ZERO;
        BigDecimal dudaSte = BigDecimal.ZERO;
        BigDecimal otros = BigDecimal.ZERO;
        BigDecimal h17 = BigDecimal.ZERO;
        try {
            var limites = locator.findRequired("LIMITES", fechaCorte);
            if (shouldSkipPoiOpen(limites, "LIMITES")) {
                throw new IllegalStateException("Insumo LIMITES muy grande para POI en modo seguro");
            }
            try (Workbook wb = WorkbookFactory.create(limites.toFile(), null, true)) {
                    Sheet aios = wb.getSheet("AIOS");
                total1 = num(aios, "AB4", null);
                deudaGobB4 = num(aios, "B4", null);
                dudaG = num(aios, "C4", null);
                dudaEf = num(aios, "E4", null);
                dudaNf = num(aios, "G4", null);
                dudaAc = num(aios, "I4", null);
                dudaF = num(aios, "K4", null);
                var ge = num(aios, "O4", null);
                var efe = num(aios, "Q4", null);
                var nfe = num(aios, "S4", null);
                var ace = num(aios, "U4", null);
                var fe = num(aios, "W4", null);
                var ste = num(aios, "Y4", null);
                dudaGe = ge;
                dudaEfe = efe;
                dudaNfe = nfe;
                dudaAce = ace;
                dudaFe = fe;
                dudaSte = ste;
                otros = num(aios, "AA4", null);
                h17 = ge.add(efe).add(nfe).add(ace).add(fe).add(ste);
            }
        } catch (OutOfMemoryError oom) {
            log.warn("OOM leyendo LIMITES; columnas 6-13 del mensual se dejarán en 0");
        } catch (Exception ignored) {
            log.warn("Insumo LIMITES no encontrado; columnas 6-13 del mensual se dejarán en 0");
        }
        log.info("Lectura LIMITES completada para fechaCorte={}", fechaCorte);

        String mes = fechaCorte.getMonth().getDisplayName(TextStyle.SHORT, new Locale("es", "CO")).replace(".", "").toLowerCase();
        String textoFecha = mes + "-" + String.format("%02d", fechaCorte.getYear() % 100);

        BigDecimal afiliados = afiliadosQuery;
        BigDecimal trm = trmService.obtener(fechaCorte);
        SeriesEconomicasService.SeriesEconomicas seriesEconomicas = seriesEconomicasService.leer(fechaCorte);
        BalanceContableQueryService.BalanceContable balanceContable = balanceContableQueryService.leer(fechaCorte);
        BigDecimal pea = seriesEconomicas.pea();
        BigDecimal deudaG = seriesEconomicas.deudaGubernamental();
        BigDecimal activosCuentas = balanceContable.activoMmCop();
        BigDecimal pasivosCuentas = balanceContable.pasivoMmCop();
        BigDecimal patrimonioCuentas = balanceContable.patrimonioMmCop();
        BigDecimal pibSemestral = seriesEconomicas.pibSemestral();
        Formato495QueryService.PensionadosResumen pensionados = formato495QueryService.leerResumen(fechaCorte);
        totalPen = pensionados.total();
        totalInv = pensionados.invalidez();
        totalVej = pensionados.vejez();
        totalSob = pensionados.sobrevivencia();
        log.info("Consulta Teradata Formato 495 completada para fechaCorte={}", fechaCorte);
        log.info("TRM seleccionada para fechaCorte={}: {}", fechaCorte, trm);

        return new MensualData(
                textoFecha,
                hombres,
                mujeres,
                afiliadosMenor30,
                afiliados30a44,
                afiliados45a59,
                afiliadosMayor60,
                afiliados,
                afiliadosActivos,
                aportantes,
                aportantesSemestral,
                traspasosSistema,
                vrFondo,
                trm,
                tmpNominal1,
                tmpReal1,
                consFdosAdmon,
                porcVrFondo,
                total1,
                dudaG,
                dudaEf,
                dudaNf,
                dudaAc,
                dudaF,
                h17,
                otros,
                dudaGe,
                dudaEfe,
                dudaNfe,
                dudaAce,
                dudaFe,
                dudaSte,
                pea,
                deudaG,
                pibSemestral,
                trm.signum() == 0 ? BigDecimal.ZERO : smColombiaCop.divide(trm, 8, RoundingMode.HALF_UP),
                totalPen,
                totalInv,
                totalVej,
                totalSob,
                fondoSistemaJ14,
                deudaGobB4,
                activosCuentas,
                pasivosCuentas,
                patrimonioCuentas,
                balanceContable.numeroAdministradorasVigentes()
        );
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

    private String findSheetPathByName(ZipFile zip, String sheetName) throws Exception {
        var dbf = DocumentBuilderFactory.newInstance();
        dbf.setNamespaceAware(true);
        var db = dbf.newDocumentBuilder();
        var wb = db.parse(zip.getInputStream(zip.getEntry("xl/workbook.xml")));
        var sheets = wb.getElementsByTagNameNS("http://schemas.openxmlformats.org/spreadsheetml/2006/main", "sheet");
        String rid = null;
        for (int i = 0; i < sheets.getLength(); i++) {
            var n = sheets.item(i);
            var name = n.getAttributes().getNamedItem("name");
            if (name != null && sheetName.equalsIgnoreCase(name.getNodeValue())) {
                var idAttr = n.getAttributes().getNamedItemNS("http://schemas.openxmlformats.org/officeDocument/2006/relationships", "id");
                if (idAttr != null) { rid = idAttr.getNodeValue(); break; }
            }
        }
        if (rid == null) return null;
        var rels = db.parse(zip.getInputStream(zip.getEntry("xl/_rels/workbook.xml.rels")));
        var relNodes = rels.getElementsByTagNameNS("http://schemas.openxmlformats.org/package/2006/relationships", "Relationship");
        for (int i = 0; i < relNodes.getLength(); i++) {
            var n = relNodes.item(i);
            var id = n.getAttributes().getNamedItem("Id");
            if (id != null && rid.equals(id.getNodeValue())) {
                var target = n.getAttributes().getNamedItem("Target");
                if (target != null) return "xl/" + target.getNodeValue().replace("\\", "/");
            }
        }
        return null;
    }

    private SexTotals readAfiliadosFromDataXml(Path file491, LocalDate fechaCorte) {
        double fechaObjetivo = DateUtil.getExcelDate(java.sql.Date.valueOf(fechaCorte));
        BigDecimal hombres = BigDecimal.ZERO;
        BigDecimal mujeres = BigDecimal.ZERO;
        try (ZipFile zip = new ZipFile(file491.toFile())) {
            String sheetPath = findSheetPathByName(zip, "Data");
            if (sheetPath == null) {
                return new SexTotals(BigDecimal.ZERO, BigDecimal.ZERO);
            }
            XMLInputFactory factory = XMLInputFactory.newFactory();
            try (InputStream is = zip.getInputStream(zip.getEntry(sheetPath))) {
                XMLStreamReader xr = factory.createXMLStreamReader(is);
                Double e = null, k = null, dv = null, dy = null;
                String cellRef = null;
                boolean inV = false;
                while (xr.hasNext()) {
                    int ev = xr.next();
                    if (ev == XMLStreamConstants.START_ELEMENT) {
                        String name = xr.getLocalName();
                        if ("row".equals(name)) {
                            e = k = dv = dy = null;
                        } else if ("c".equals(name)) {
                            cellRef = xr.getAttributeValue(null, "r");
                        } else if ("v".equals(name)) {
                            inV = true;
                        }
                    } else if (ev == XMLStreamConstants.CHARACTERS && inV && cellRef != null) {
                        String t = xr.getText();
                        if (t != null && !t.isBlank()) {
                            try {
                                double n = Double.parseDouble(t.trim());
                                if (cellRef.startsWith("E")) e = n;
                                else if (cellRef.startsWith("K")) k = n;
                                else if (cellRef.startsWith("DV")) dv = n;
                                else if (cellRef.startsWith("DY")) dy = n;
                            } catch (NumberFormatException ignored) {
                            }
                        }
                    } else if (ev == XMLStreamConstants.END_ELEMENT) {
                        String name = xr.getLocalName();
                        if ("v".equals(name)) inV = false;
                        if ("row".equals(name)) {
                            if (e != null && k != null && Math.abs(e - fechaObjetivo) < 0.00001d && Math.abs(k - 999d) < 0.00001d) {
                                if (dv != null) hombres = hombres.add(BigDecimal.valueOf(dv));
                                if (dy != null) mujeres = mujeres.add(BigDecimal.valueOf(dy));
                            }
                        }
                    }
                }
                xr.close();
            }
        } catch (Exception ignored) {
            return new SexTotals(BigDecimal.ZERO, BigDecimal.ZERO);
        }
        return new SexTotals(hombres, mujeres);
    }

    private record SexTotals(BigDecimal hombres, BigDecimal mujeres) {}

    private BigDecimal parseNumber(Cell cell, DataFormatter formatter) {
        if (cell == null) return BigDecimal.ZERO;
        try {
            if (cell.getCellType() == CellType.NUMERIC) {
                return BigDecimal.valueOf(cell.getNumericCellValue());
            }
            String text = formatter.formatCellValue(cell);
            if (text == null || text.isBlank()) return BigDecimal.ZERO;
            String normalized = text.replace(".", "").replace(",", ".").replace(" ", "").trim();
            return new BigDecimal(normalized);
        } catch (Exception e) {
            return BigDecimal.ZERO;
        }
    }

    private BigDecimal readNumericCellFromSheetXml(Path file, String sheetName, String cellRefWanted) {
        try (ZipFile zip = new ZipFile(file.toFile())) {
            String sheetPath = findSheetPathByName(zip, sheetName);
            if (sheetPath == null) {
                return BigDecimal.ZERO;
            }
            XMLInputFactory factory = XMLInputFactory.newFactory();
            try (InputStream is = zip.getInputStream(zip.getEntry(sheetPath))) {
                XMLStreamReader xr = factory.createXMLStreamReader(is);
                String cellRef = null;
                boolean inV = false;
                while (xr.hasNext()) {
                    int ev = xr.next();
                    if (ev == XMLStreamConstants.START_ELEMENT) {
                        String name = xr.getLocalName();
                        if ("c".equals(name)) {
                            cellRef = xr.getAttributeValue(null, "r");
                        } else if ("v".equals(name) && cellRefWanted.equals(cellRef)) {
                            inV = true;
                        }
                    } else if (ev == XMLStreamConstants.CHARACTERS && inV) {
                        String t = xr.getText();
                        if (t != null && !t.isBlank()) {
                            try {
                                return BigDecimal.valueOf(Double.parseDouble(t.trim()));
                            } catch (NumberFormatException ignored) {
                                return BigDecimal.ZERO;
                            }
                        }
                    } else if (ev == XMLStreamConstants.END_ELEMENT) {
                        if ("v".equals(xr.getLocalName()) && inV) {
                            inV = false;
                        }
                    }
                }
                xr.close();
            }
        } catch (Exception ignored) {
            return BigDecimal.ZERO;
        }
        return BigDecimal.ZERO;
    }

    private boolean shouldSkipPoiOpen(Path file, String tag) {
        try {
            long bytes = Files.size(file);
            int maxMb = properties.maxPoiFileMb() == null ? 40 : properties.maxPoiFileMb();
            long maxBytes = maxMb * 1024L * 1024L;
            if (bytes > maxBytes) {
                log.warn("{} no se abrirá con POI ({} MB > {} MB configurados)", tag, bytes / (1024 * 1024), maxMb);
                return true;
            }
            return false;
        } catch (Exception e) {
            return false;
        }
    }

    private BigDecimal num(Sheet sheet, String a1, FormulaEvaluator evaluator) {
        CellReference ref = new CellReference(a1);
        return num(sheet, ref.getRow() + 1, ref.getCol() + 1, evaluator);
    }

    private BigDecimal num(Sheet sheet, int row1Based, int col1Based, FormulaEvaluator evaluator) {
        Row row = sheet.getRow(row1Based - 1);
        if (row == null) return BigDecimal.ZERO;
        Cell cell = row.getCell(col1Based - 1);
        if (cell == null) return BigDecimal.ZERO;
        try {
            if (cell.getCellType() == CellType.FORMULA) {
                if (evaluator != null) {
                    CellValue v = evaluator.evaluate(cell);
                    if (v != null && v.getCellType() == CellType.NUMERIC) return BigDecimal.valueOf(v.getNumberValue());
                    return BigDecimal.ZERO;
                }
                return BigDecimal.valueOf(cell.getNumericCellValue());
            }
            if (cell.getCellType() == CellType.NUMERIC) return BigDecimal.valueOf(cell.getNumericCellValue());
            if (cell.getCellType() == CellType.STRING) return new BigDecimal(cell.getStringCellValue().trim());
        } catch (Exception ignored) {
        }
        return BigDecimal.ZERO;
    }
}
