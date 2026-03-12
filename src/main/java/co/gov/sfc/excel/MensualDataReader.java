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

    public MensualDataReader(InsumosLocator locator, AiosProperties properties) {
        this.locator = locator;
        this.properties = properties;
        // Evitar asignaciones gigantes en POI que pueden terminar en OOM con archivos grandes.
        // 100 MB es suficiente para los insumos actuales y más conservador en memoria.
        IOUtils.setByteArrayMaxOverride(100_000_000);
    }

    public MensualData read(LocalDate fechaCorte) {
        log.info("Iniciando lectura de insumos para fechaCorte={}", fechaCorte);

        BigDecimal hombres = BigDecimal.ZERO;
        BigDecimal mujeres = BigDecimal.ZERO;
        BigDecimal aportantes = BigDecimal.ZERO;
        BigDecimal consFdosAdmon = BigDecimal.ZERO;

        var file491 = locator.findRequired("491", fechaCorte);
        var file493 = locator.findRequired("493", fechaCorte);
        boolean macroRecalc = Boolean.TRUE.equals(properties.macroRecalc491493());
        BigDecimal traspasosSistema = BigDecimal.ZERO;

        if (macroRecalc) {
            try (Workbook wb = WorkbookFactory.create(file491.toFile(), null, true)) {
                Sheet informe = wb.getSheet("informe de prensa");
                Sheet multifondos = wb.getSheet("multifondos");
                FormulaEvaluator evaluator = wb.getCreationHelper().createFormulaEvaluator();
                setDate(informe, "C3", fechaCorte);
                setDate(multifondos, "C4", fechaCorte);
                evaluator.clearAllCachedResultValues();
                hombres = num(informe, "C11", evaluator);
                mujeres = num(informe, "D11", evaluator);
                aportantes = num(multifondos, "E25", evaluator);
                var j8 = num(multifondos, "J8", evaluator);
                var j9 = num(multifondos, "J9", evaluator);
                var j12 = num(multifondos, "J12", evaluator);
                consFdosAdmon = j12.signum() == 0 ? BigDecimal.ZERO : j8.add(j9).divide(j12, 8, RoundingMode.HALF_UP).multiply(BigDecimal.valueOf(100));
            } catch (OutOfMemoryError oom) {
                log.warn("OOM en recálculo macro 491; se usa modo seguro XML cacheado");
                macroRecalc = false;
            } catch (Exception e) {
                throw new IllegalStateException("Error leyendo Formato 491", e);
            }

            if (macroRecalc) {
                try (Workbook wb = WorkbookFactory.create(file493.toFile(), null, true)) {
                    FormulaEvaluator evaluator = wb.getCreationHelper().createFormulaEvaluator();
                    Sheet tras = wb.getSheet("Traslados Entre AFP");
                    if (tras == null) {
                        throw new IllegalStateException("No existe hoja 'Traslados Entre AFP' en Formato 493");
                    }
                    setDate(tras, "B11", fechaCorte);
                    setNumeric(tras, "D4", 99);
                    evaluator.clearAllCachedResultValues();
                    traspasosSistema = num(tras, "BQ11", evaluator);
                } catch (OutOfMemoryError oom) {
                    log.warn("OOM en recálculo macro 493; se usa modo seguro XML cacheado");
                    macroRecalc = false;
                } catch (Exception e) {
                    log.warn("No fue posible leer Formato 493 con recálculo; se intenta modo seguro. Causa: {}", e.getMessage());
                    macroRecalc = false;
                }
            }
        }

        if (!macroRecalc) {
            try {
                hombres = readNumericCellFromSheetXml(file491, "informe de prensa", "C11");
                mujeres = readNumericCellFromSheetXml(file491, "informe de prensa", "D11");
                aportantes = readNumericCellFromSheetXml(file491, "multifondos", "E25");
                var j8 = readNumericCellFromSheetXml(file491, "multifondos", "J8");
                var j9 = readNumericCellFromSheetXml(file491, "multifondos", "J9");
                var j12 = readNumericCellFromSheetXml(file491, "multifondos", "J12");
                consFdosAdmon = j12.signum() == 0 ? BigDecimal.ZERO : j8.add(j9).divide(j12, 8, RoundingMode.HALF_UP).multiply(BigDecimal.valueOf(100));
                log.info("Lectura 491 en modo seguro XML cacheado para fechaCorte={}", fechaCorte);
            } catch (Exception e) {
                throw new IllegalStateException("Error leyendo Formato 491", e);
            }

            try {
                traspasosSistema = readNumericCellFromSheetXml(file493, "Traslados Entre AFP", "BQ11");
                log.info("Lectura 493 en modo seguro XML cacheado para fechaCorte={}", fechaCorte);
            } catch (Exception e) {
                log.warn("No fue posible leer Formato 493; se usará 0 en traspasos_sistema. Causa: {}", e.getMessage());
            }
        }
        log.info("Lectura Formato 491 completada para fechaCorte={}", fechaCorte);
        log.info("Lectura Formato 493 completada para fechaCorte={}", fechaCorte);

        BigDecimal tmpReal1;
        BigDecimal tmpNominal1;
        var rentFile = locator.findRequired("Rent_Vr_Uni_Moderado", fechaCorte);
        try {
            RentResult rent = readRentabilidadDesdeXml(rentFile, fechaCorte);
            tmpNominal1 = rent.nominal();
            tmpReal1 = rent.real();
            log.info("Rentabilidad moderado (stream xml) con fechaCorte={}: D11(nominal)={}, D10(real)={}",
                    fechaCorte, tmpNominal1, tmpReal1);
        } catch (Exception fastError) {
            log.warn("Lectura streaming de rentabilidad falló (se intenta fallback POI): {}", fastError.getMessage());
            try (Workbook wb = WorkbookFactory.create(rentFile.toFile(), null, true)) {
                Sheet consolidado = getSheetIgnoreCase(wb, "Consolidado");
                if (consolidado == null) consolidado = wb.getSheetAt(0);
                LocalDate fechaInicial = fechaCorte.minusYears(1);
                setDate(consolidado, "D5", fechaCorte);
                setDate(consolidado, "D4", fechaInicial);
                tmpNominal1 = readRentabilidadNominal(consolidado, fechaInicial, fechaCorte);
                tmpReal1 = readRentabilidadReal(consolidado, fechaCorte);
            } catch (Exception fallbackError) {
                throw new IllegalStateException("Error leyendo rentabilidad moderado", fallbackError);
            }
        }
        log.info("Lectura rentabilidad completada para fechaCorte={}", fechaCorte);

        BigDecimal vrFondo = BigDecimal.ZERO;
        BigDecimal porcVrFondo = BigDecimal.ZERO;
        var sistemaTotal = locator.findRequired("SISTEMA TOTAL", fechaCorte);
        try {
            if (shouldSkipPoiOpen(sistemaTotal, "SISTEMA TOTAL")) {
                throw new IllegalStateException("Insumo muy grande para POI en modo seguro");
            }
            try (Workbook wb = WorkbookFactory.create(sistemaTotal.toFile(), null, true)) {
            Sheet ws = wb.getSheet("restot");
            int cSistema = findHeaderCol(ws, "SISTEMA");
            int cProt = findHeaderCol(ws, "PROTECCION");
            int cPorv = findHeaderCol(ws, "PORVENIR");
            int row = findMaxRow(ws, cSistema + 1, null);
            vrFondo = num(ws, row, cSistema + 1, null).divide(BigDecimal.valueOf(1000), 8, RoundingMode.HALF_UP);
            var prot = num(ws, row, cProt + 1, null);
            var porv = num(ws, row, cPorv + 1, null);
            if (vrFondo.signum() != 0) {
                porcVrFondo = prot.add(porv).divide(vrFondo, 8, RoundingMode.HALF_UP).divide(BigDecimal.TEN, 8, RoundingMode.HALF_UP);
            }
            }
        } catch (OutOfMemoryError oom) {
            log.warn("OOM leyendo SISTEMA TOTAL; se usarán ceros para este bloque");
        } catch (Exception e) {
            log.warn("No fue posible leer SISTEMA TOTAL: {}", e.getMessage());
        }
        log.info("Lectura SISTEMA TOTAL completada para fechaCorte={}", fechaCorte);

        BigDecimal total1 = BigDecimal.ZERO;
        BigDecimal dudaG = BigDecimal.ZERO;
        BigDecimal dudaEf = BigDecimal.ZERO;
        BigDecimal dudaNf = BigDecimal.ZERO;
        BigDecimal dudaAc = BigDecimal.ZERO;
        BigDecimal dudaF = BigDecimal.ZERO;
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

        BigDecimal trm = readTrmFromSeries(fechaCorte);
        log.info("TRM seleccionada para fechaCorte={}: {}", fechaCorte, trm);

        return new MensualData(
                textoFecha,
                hombres.add(mujeres),
                aportantes,
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
                otros
        );
    }


    private BigDecimal readTrmFromSeries(LocalDate fechaCorte) {
        try {
            var seriesFile = locator.findRequired("PIB_PEA_TRM_DG", fechaCorte);
            if (shouldSkipPoiOpen(seriesFile, "PIB_PEA_TRM_DG")) {
                return BigDecimal.ONE;
            }
            try (Workbook wb = WorkbookFactory.create(seriesFile.toFile(), null, true)) {
                Sheet sheet = wb.getSheet("Hoja1");
                if (sheet == null) {
                    sheet = wb.getSheetAt(0);
                }
                BigDecimal trm = BigDecimal.ONE;
                LocalDate mejorFecha = LocalDate.MIN;
                for (Row row : sheet) {
                    LocalDate fecha = cellAsDate(row.getCell(1)); // Columna B
                    if (fecha == null || fecha.isAfter(fechaCorte)) {
                        continue;
                    }
                    BigDecimal valor = num(sheet, row.getRowNum() + 1, 3, null); // Columna C
                    if (!fecha.isBefore(mejorFecha)) {
                        mejorFecha = fecha;
                        trm = valor;
                    }
                }
                return trm.signum() == 0 ? BigDecimal.ONE : trm;
            }
        } catch (Exception e) {
            log.warn("No se pudo leer TRM desde series PIB_PEA_TRM_DG: {}", e.getMessage());
            return BigDecimal.ONE;
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

    private BigDecimal readRentabilidadNominal(Sheet consolidado, LocalDate fechaInicial, LocalDate fechaFinal) {
        BigDecimal valorInicial = lookupByDate(consolidado, 5, fechaInicial, true);
        BigDecimal valorFinal = lookupByDate(consolidado, 5, fechaFinal, true);
        if (valorInicial == null || valorFinal == null || valorInicial.signum() == 0) {
            return BigDecimal.ZERO;
        }
        double dias = Math.max(1d, fechaFinal.toEpochDay() - fechaInicial.toEpochDay());
        double nominal = Math.pow(valorFinal.doubleValue() / valorInicial.doubleValue(), 365d / dias) - 1d;
        return BigDecimal.valueOf(nominal);
    }

    private BigDecimal lookupByDate(Sheet sheet, int valueCol1Based, LocalDate target, boolean allowPrevious) {
        double objetivo = DateUtil.getExcelDate(java.sql.Date.valueOf(target));
        BigDecimal exacta = null;
        BigDecimal anterior = null;
        double fechaAnterior = Double.NEGATIVE_INFINITY;
        int last = sheet.getLastRowNum() + 1;
        for (int r = 14; r <= last; r++) {
            BigDecimal fecha = num(sheet, r, 1, null);
            if (fecha.signum() == 0) continue;
            double excelDate = fecha.doubleValue();
            BigDecimal valor = num(sheet, r, valueCol1Based, null);
            if (Math.abs(excelDate - objetivo) < 0.00001d && valor.signum() != 0) {
                exacta = valor;
                break;
            }
            if (allowPrevious && excelDate <= objetivo && excelDate > fechaAnterior && valor.signum() != 0) {
                fechaAnterior = excelDate;
                anterior = valor;
            }
        }
        return exacta != null ? exacta : anterior;
    }

    private BigDecimal readRentabilidadReal(Sheet consolidado, LocalDate fechaCorte) {
        // En macro VBA D10 equivale a BUSCARV(fecha_final, A:I, 9, FALSO).
        // POI puede devolver 0 cuando D10 evalúa error por dependencias externas/caché,
        // por eso se hace el lookup explícito sobre la tabla base.
        // Importante: se usan valores cacheados (sin evaluator) para evitar evaluar miles de fórmulas y disparar uso de heap.
        double objetivo = DateUtil.getExcelDate(java.sql.Date.valueOf(fechaCorte));
        BigDecimal exacta = null;
        BigDecimal anterior = null;
        double fechaAnterior = Double.NEGATIVE_INFINITY;

        int last = consolidado.getLastRowNum() + 1;
        for (int r = 14; r <= last; r++) {
            BigDecimal fecha = num(consolidado, r, 1, null);
            if (fecha.signum() == 0) {
                continue;
            }
            double excelDate = fecha.doubleValue();
            BigDecimal real = num(consolidado, r, 9, null);
            if (Math.abs(excelDate - objetivo) < 0.00001d && real.signum() != 0) {
                exacta = real;
                break;
            }
            if (excelDate <= objetivo && excelDate > fechaAnterior && real.signum() != 0) {
                fechaAnterior = excelDate;
                anterior = real;
            }
        }

        if (exacta != null) return exacta;
        if (anterior != null) {
            log.info("Rentabilidad real D10 sin match exacto para {}. Se usa fecha hábil anterior (excelDate={})", fechaCorte, fechaAnterior);
            return anterior;
        }
        return num(consolidado, "D10", null);
    }


    private RentResult readRentabilidadDesdeXml(Path rentFile, LocalDate fechaCorte) throws Exception {
        double objetivoFinal = DateUtil.getExcelDate(java.sql.Date.valueOf(fechaCorte));
        double objetivoInicial = DateUtil.getExcelDate(java.sql.Date.valueOf(fechaCorte.minusYears(1)));

        try (ZipFile zip = new ZipFile(rentFile.toFile())) {
            String sheetPath = findSheetPathByName(zip, "Consolidado");
            if (sheetPath == null) throw new IllegalStateException("No se encontró hoja Consolidado en workbook.xml");

            BigDecimal eIni = null, eFin = null, eIniPrev = null, eFinPrev = null;
            double eIniPrevDate = Double.NEGATIVE_INFINITY, eFinPrevDate = Double.NEGATIVE_INFINITY;
            BigDecimal iFin = null, iFinPrev = null;
            double iFinPrevDate = Double.NEGATIVE_INFINITY;

            XMLInputFactory factory = XMLInputFactory.newFactory();
            try (InputStream is = zip.getInputStream(zip.getEntry(sheetPath))) {
                XMLStreamReader xr = factory.createXMLStreamReader(is);
                int rowNum = -1;
                Double aVal = null, eVal = null, iVal = null;
                String cellRef = null;
                boolean inV = false;
                while (xr.hasNext()) {
                    int ev = xr.next();
                    if (ev == XMLStreamConstants.START_ELEMENT) {
                        String name = xr.getLocalName();
                        if ("row".equals(name)) {
                            String r = xr.getAttributeValue(null, "r");
                            rowNum = r == null ? -1 : Integer.parseInt(r);
                            aVal = eVal = iVal = null;
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
                                if (cellRef.startsWith("A")) aVal = n;
                                else if (cellRef.startsWith("E")) eVal = n;
                                else if (cellRef.startsWith("I")) iVal = n;
                            } catch (NumberFormatException ignored) {
                            }
                        }
                    } else if (ev == XMLStreamConstants.END_ELEMENT) {
                        String name = xr.getLocalName();
                        if ("v".equals(name)) inV = false;
                        if ("row".equals(name) && rowNum >= 14 && aVal != null) {
                            double d = aVal;
                            if (eVal != null && eVal != 0d) {
                                if (Math.abs(d - objetivoInicial) < 0.00001d) eIni = BigDecimal.valueOf(eVal);
                                if (Math.abs(d - objetivoFinal) < 0.00001d) eFin = BigDecimal.valueOf(eVal);
                                if (d <= objetivoInicial && d > eIniPrevDate) { eIniPrevDate = d; eIniPrev = BigDecimal.valueOf(eVal); }
                                if (d <= objetivoFinal && d > eFinPrevDate) { eFinPrevDate = d; eFinPrev = BigDecimal.valueOf(eVal); }
                            }
                            if (iVal != null && iVal != 0d) {
                                if (Math.abs(d - objetivoFinal) < 0.00001d) iFin = BigDecimal.valueOf(iVal);
                                if (d <= objetivoFinal && d > iFinPrevDate) { iFinPrevDate = d; iFinPrev = BigDecimal.valueOf(iVal); }
                            }
                        }
                    }
                }
                xr.close();
            }

            BigDecimal vi = eIni != null ? eIni : eIniPrev;
            BigDecimal vf = eFin != null ? eFin : eFinPrev;
            BigDecimal nominal = BigDecimal.ZERO;
            if (vi != null && vf != null && vi.signum() != 0) {
                double dias = Math.max(1d, fechaCorte.toEpochDay() - fechaCorte.minusYears(1).toEpochDay());
                nominal = BigDecimal.valueOf(Math.pow(vf.doubleValue() / vi.doubleValue(), 365d / dias) - 1d);
            }
            BigDecimal real = iFin != null ? iFin : iFinPrev;
            if (real == null) real = BigDecimal.ZERO;
            return new RentResult(nominal, real);
        }
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

    private record RentResult(BigDecimal nominal, BigDecimal real) {}


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

    private LocalDate cellAsDate(Cell cell) {
        if (cell == null) {
            return null;
        }
        try {
            if (cell.getCellType() == CellType.NUMERIC && DateUtil.isCellDateFormatted(cell)) {
                return cell.getLocalDateTimeCellValue().toLocalDate();
            }
            String txt = new DataFormatter().formatCellValue(cell);
            if (txt == null || txt.isBlank()) {
                return null;
            }
            return java.time.LocalDate.parse(txt, java.time.format.DateTimeFormatter.ofPattern("d/M/yyyy"));
        } catch (Exception e) {
            return null;
        }
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

    private int findHeaderCol(Sheet sheet, String header) {
        for (Row row : sheet) {
            for (Cell cell : row) {
                if (cell.getCellType() == CellType.STRING && cell.getStringCellValue().toUpperCase().contains(header.toUpperCase())) {
                    return cell.getColumnIndex();
                }
            }
        }
        throw new IllegalArgumentException("No header " + header);
    }

    private int findMaxRow(Sheet sheet, int col1Based, FormulaEvaluator evaluator) {
        int bestRow = -1;
        BigDecimal max = BigDecimal.valueOf(-1);
        for (Row row : sheet) {
            BigDecimal v = num(sheet, row.getRowNum() + 1, col1Based, evaluator);
            if (v.compareTo(max) > 0) {
                max = v;
                bestRow = row.getRowNum() + 1;
            }
        }
        if (bestRow < 1) throw new IllegalArgumentException("No data row");
        return bestRow;
    }


    private BigDecimal sumRange(Sheet sheet, FormulaEvaluator evaluator, int rowStart, int rowEnd, int col1Based) {
        BigDecimal total = BigDecimal.ZERO;
        for (int r = rowStart; r <= rowEnd; r++) {
            total = total.add(num(sheet, r, col1Based, evaluator));
        }
        return total;
    }

    private void setDate(Sheet sheet, String a1, LocalDate value) {
        CellReference ref = new CellReference(a1);
        Row row = sheet.getRow(ref.getRow());
        if (row == null) row = sheet.createRow(ref.getRow());
        Cell cell = row.getCell(ref.getCol());
        if (cell == null) cell = row.createCell(ref.getCol());
        cell.setCellValue(java.sql.Date.valueOf(value));
    }

    private void setNumeric(Sheet sheet, String a1, double value) {
        CellReference ref = new CellReference(a1);
        Row row = sheet.getRow(ref.getRow());
        if (row == null) row = sheet.createRow(ref.getRow());
        Cell cell = row.getCell(ref.getCol());
        if (cell == null) cell = row.createCell(ref.getCol());
        cell.setCellValue(value);
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
