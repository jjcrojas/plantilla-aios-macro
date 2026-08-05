package co.gov.sfc.excel;

import co.gov.sfc.insumos.InsumosLocator;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellReference;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Component;

import java.math.BigDecimal;
import java.nio.file.Files;
import java.nio.file.Path;
import java.text.Normalizer;
import java.time.LocalDate;
import java.time.format.TextStyle;
import java.util.Arrays;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Locale;
import java.util.Map;
import java.util.Set;

@Component
public class TrimestralDataReader {

    private static final Logger log = LoggerFactory.getLogger(TrimestralDataReader.class);
    private final MensualDataReader mensualDataReader;
    private final InsumosLocator locator;
    private final Formato491QueryService formato491QueryService;
    private final Formato493QueryService formato493QueryService;
    private final Formato136QueryService formato136QueryService;

    public TrimestralDataReader(MensualDataReader mensualDataReader, InsumosLocator locator, Formato491QueryService formato491QueryService, Formato493QueryService formato493QueryService, Formato136QueryService formato136QueryService) {
        this.mensualDataReader = mensualDataReader;
        this.locator = locator;
        this.formato491QueryService = formato491QueryService;
        this.formato493QueryService = formato493QueryService;
        this.formato136QueryService = formato136QueryService;
    }

    public TrimestralData read(LocalDate fechaCorte) {
        return read(fechaCorte, mensualDataReader.read(fechaCorte));
    }

    public TrimestralData read(LocalDate fechaCorte, MensualData mensual) {

        log.info("Afiliados trimestrales del Formato 491 se consultarán en Teradata; no se requiere archivo Excel local 491 para fechaCorte={}", fechaCorte);
        Map<String, BigDecimal> afiliados = formato491QueryService.leerAfiliadosTrimestralPorFondo(fechaCorte);
        Map<String, BigDecimal> aportantes = formato491QueryService.leerAportantesPorEntidad(fechaCorte);
        Map<String, BigDecimal> traspasos = formato493QueryService.leerTraspasosPorEntidad(fechaCorte);
        Map<String, BigDecimal> colombiaUsd = readColombiaUsd(fechaCorte, mensual.trm());
        Map<String, BigDecimal> gastosUsd = readGastosUsd(fechaCorte, mensual.trm());
        Map<String, BigDecimal> comisionesPct = readComisiones(fechaCorte);
        Map<String, BigDecimal> rentNominalPct = new HashMap<>();
        Map<String, BigDecimal> rentRealPct = new HashMap<>();
        readRentabilidad(fechaCorte, rentNominalPct, rentRealPct);

        String etiquetaFecha = fechaCorte.getMonth().getDisplayName(TextStyle.SHORT, new Locale("es", "CO"))
                .replace(".", "")
                .toLowerCase() + "-" + String.format("%02d", fechaCorte.getYear() % 100);

        return new TrimestralData(etiquetaFecha, afiliados, aportantes, traspasos, colombiaUsd, gastosUsd, comisionesPct, rentNominalPct, rentRealPct);
    }

    private Map<String, BigDecimal> readColombiaUsd(LocalDate fechaCorte, BigDecimal trm) {
        Map<String, BigDecimal> saldosMillonesCop = formato136QueryService.leerColombiaPorFondoEntidad(fechaCorte);
        Map<String, BigDecimal> out = new HashMap<>();
        saldosMillonesCop.forEach((key, value) -> out.put(key, safeDivide(value, trm)));
        log.info("Colombia trimestral: saldos Formato 136 consultados en Teradata para fechaCorte={} y convertidos con TRM={} valores={}",
                fechaCorte, trm, out);
        return out;
    }

    private void readBalanceTo(Map<String, BigDecimal> out, Path dir, String name, String pref, boolean allowAlt, BigDecimal trm) throws Exception {
        Path f = findInDirContains(dir, name);
        if (f == null) return;
        try (Workbook wb = WorkbookFactory.create(f.toFile(), null, true)) {
            Sheet ws = getSheetIgnoreCase(wb, "restot");
            if (ws == null) ws = wb.getSheetAt(0);
            RowData r = readBalanceRow(ws, allowAlt);
            out.put(pref + "_colf", safeDivide(r.colf, trm));
            out.put(pref + "_porv", safeDivide(r.porv, trm));
            out.put(pref + "_prot", safeDivide(r.prot, trm));
            out.put(pref + "_sk", safeDivide(r.skan, trm));
            if (allowAlt) out.put(pref + "_alt", safeDivide(r.alt, trm));
        }
    }

    private Map<String, BigDecimal> readGastosUsd(LocalDate fechaCorte, BigDecimal trm) {
        Map<String, BigDecimal> out = new HashMap<>();
        try {
            Path plantilla = findPlantillaAiosFile(fechaCorte);
            log.info("Gastos trimestrales: leyendo plantilla {}", plantilla.toAbsolutePath());
            try (Workbook wb = WorkbookFactory.create(plantilla.toFile(), null, true)) {
                Sheet baseAnual = getSheetIgnoreCase(wb, "base anual");
                if (baseAnual == null) {
                    log.warn("No se encontró la hoja 'base anual' en {}", plantilla.getFileName());
                    return out;
                }
                FormulaEvaluator evaluator = wb.getCreationHelper().createFormulaEvaluator();
                LocalDate fechaBase = fechaBusquedaGastos(fechaCorte);
                int serialFecha = (int) Math.round(DateUtil.getExcelDate(java.sql.Date.valueOf(fechaBase)));
                log.info("Gastos trimestrales: fechaCorte={}, fechaBaseBusqueda={}, serialExcel={}, TRM={}", fechaCorte, fechaBase, serialFecha, trm);

                putGastoUsd(out, "prot", "proteccion", baseAnual, evaluator, serialFecha, trm);
                putGastoUsd(out, "porv", "porvenir", baseAnual, evaluator, serialFecha, trm);
                putGastoUsd(out, "sk", "skandia", baseAnual, evaluator, serialFecha, trm);
                putGastoUsd(out, "colf", "colfondos", baseAnual, evaluator, serialFecha, trm);
            }
        } catch (Exception e) {
            log.warn("No se pudo leer gastos trimestrales: {}", e.getMessage());
        }
        return out;
    }

    private LocalDate fechaBusquedaGastos(LocalDate fechaCorte) {
        return fechaCorte.withDayOfMonth(1);
    }

    private Path findPlantillaAiosFile(LocalDate fechaCorte) {
        try {
            return locator.findRequired("Plantilla AIOS-probable", fechaCorte);
        } catch (Exception ignore1) {
            try {
                return locator.findRequired("Plantilla_AIOS", fechaCorte);
            } catch (Exception ignore2) {
            Path repoPath = Path.of("plantillas", "Plantilla AIOS-probable.xlsm");
            if (Files.isRegularFile(repoPath)) return repoPath;
            Path localPath = Path.of("Plantilla AIOS-probable.xlsm");
            if (Files.isRegularFile(localPath)) return localPath;
            throw new IllegalStateException("No se encontró Plantilla AIOS-probable.xlsm para lectura de gastos.");
            }
        }
    }

    private void putGastoUsd(Map<String, BigDecimal> out, String key, String administradora, Sheet baseAnual, FormulaEvaluator evaluator, int serialFecha, BigDecimal trm) {
        BigDecimal gastoMillonesCop = gastoNetoCop(baseAnual, evaluator, administradora, serialFecha);
        BigDecimal gastoUsd = safeDivide(gastoMillonesCop, trm);
        out.put(key, gastoUsd);
        log.info("Gastos {}: neto_MCOP={} -> USD={}", administradora, gastoMillonesCop, gastoUsd);
    }

    private BigDecimal gastoNetoCop(Sheet baseAnual, FormulaEvaluator evaluator, String administradora, int serialFecha) {
        Set<String> cuentasDescuento = new HashSet<>(Arrays.asList(
                "510300", "510400", "510600", "510700", "510800", "512500", "512800", "512900", "513900"
        ));
        Set<String> cuentasObjetivo = new HashSet<>(cuentasDescuento);
        cuentasObjetivo.add("510000");

        DataFormatter fmt = new DataFormatter();
        Map<String, BigDecimal> valores = new HashMap<>();

        for (int r = 1; r <= baseAnual.getLastRowNum(); r++) {
            Row row = baseAnual.getRow(r);
            if (row == null) continue;
            String adminFila = normalize(fmt.formatCellValue(row.getCell(2), evaluator)); // col C
            int serialFila = excelSerial(row.getCell(1), evaluator); // col B
            String cuenta = normalize(fmt.formatCellValue(row.getCell(3), evaluator)).replace(".0", ""); // col D
            if (!normalize(administradora).equals(adminFila) || serialFila != serialFecha) continue;
            if (!cuentasObjetivo.contains(cuenta) || valores.containsKey(cuenta)) continue;
            valores.put(cuenta, num(row.getCell(6), null)); // columna G
            if (valores.size() == cuentasObjetivo.size()) break;
        }

        BigDecimal gasto = valores.getOrDefault("510000", BigDecimal.ZERO);
        BigDecimal descuentos = BigDecimal.ZERO;
        for (String c : cuentasDescuento) descuentos = descuentos.add(valores.getOrDefault(c, BigDecimal.ZERO));
        gasto = gasto.subtract(descuentos);

        if (!valores.containsKey("510000")) {
            log.warn("Gastos {}: no se encontró cuenta 510000 para serial {}.", administradora, serialFecha);
        }
        log.info("Gastos {} serial {}: 510000={}, descuentos={}, cuentas_encontradas={}", administradora, serialFecha, valores.getOrDefault("510000", BigDecimal.ZERO), descuentos, valores.keySet());
        return gasto.divide(BigDecimal.valueOf(1_000_000), 8, java.math.RoundingMode.HALF_UP);
    }

    private int excelSerial(Cell c, FormulaEvaluator eval) {
        if (c == null) return Integer.MIN_VALUE;
        try {
            if (c.getCellType() == CellType.NUMERIC) {
                return (int) Math.round(c.getNumericCellValue());
            }
            if (c.getCellType() == CellType.FORMULA && eval != null) {
                CellValue cv = eval.evaluate(c);
                if (cv != null && cv.getCellType() == CellType.NUMERIC) return (int) Math.round(cv.getNumberValue());
            }
            String txt = new DataFormatter().formatCellValue(c, eval).trim();
            if (txt.isBlank()) return Integer.MIN_VALUE;
            return (int) Math.round(Double.parseDouble(txt.replace(",", ".")));
        } catch (Exception e) {
            return Integer.MIN_VALUE;
        }
    }

    private Map<String, BigDecimal> readComisiones(LocalDate fechaCorte) {
        Map<String, BigDecimal> out = new HashMap<>();
        try {
            Path file = findComisionesFile(fechaCorte);
            try (Workbook wb = WorkbookFactory.create(file.toFile(), null, true)) {
                Sheet ws = getSheetIgnoreCase(wb, "COTIZACION CORTE ANUAL");
                FormulaEvaluator eval = wb.getCreationHelper().createFormulaEvaluator();
                setDate(ws, "A1", fechaCorte);
                eval.clearAllCachedResultValues();
                out.put("ska_obl", num(ws, "B1", eval).multiply(BigDecimal.valueOf(100)));
                out.put("ska_seg", num(ws, "C1", eval).multiply(BigDecimal.valueOf(100)));
                out.put("por_obl", num(ws, "F1", eval).multiply(BigDecimal.valueOf(100)));
                out.put("por_seg", num(ws, "G1", eval).multiply(BigDecimal.valueOf(100)));
                out.put("pro_obl", num(ws, "N1", eval).multiply(BigDecimal.valueOf(100)));
                out.put("pro_seg", num(ws, "O1", eval).multiply(BigDecimal.valueOf(100)));
                out.put("col_obl", num(ws, "R1", eval).multiply(BigDecimal.valueOf(100)));
                out.put("col_seg", num(ws, "S1", eval).multiply(BigDecimal.valueOf(100)));
            }
        } catch (Exception e) {
            log.warn("No se pudo leer comisiones trimestrales: {}", e.getMessage());
        }
        return out;
    }

    private void readRentabilidad(LocalDate fechaCorte, Map<String, BigDecimal> nom, Map<String, BigDecimal> real) {
        try {
            Path file = locator.findRequired("Rent_Vr_Uni_Moderado", fechaCorte);
            try (Workbook wb = WorkbookFactory.create(file.toFile(), null, true)) {
                FormulaEvaluator eval = wb.getCreationHelper().createFormulaEvaluator();
                readRentSheet(wb, eval, fechaCorte, "Colfondos", "colf", nom, real);
                readRentSheet(wb, eval, fechaCorte, "Porvenir", "porv", nom, real);
                readRentSheet(wb, eval, fechaCorte, "Protección", "prot", nom, real);
                if (getSheetIgnoreCase(wb, "Protección") == null) readRentSheet(wb, eval, fechaCorte, "Proteccion", "prot", nom, real);
                readRentSheet(wb, eval, fechaCorte, "oldmutual", "oldm", nom, real);
            }
        } catch (Exception e) {
            log.warn("No se pudo leer rentabilidad trimestral: {}", e.getMessage());
        }
    }

    private void readRentSheet(Workbook wb, FormulaEvaluator eval, LocalDate fecha, String sheetName, String key, Map<String, BigDecimal> nom, Map<String, BigDecimal> real) {
        Sheet s = getSheetIgnoreCase(wb, sheetName);
        if (s == null) return;
        setDate(s, "D5", fecha);
        setDate(s, "D4", fecha.minusYears(1));
        eval.clearAllCachedResultValues();
        real.put(key, num(s, "D10", eval).multiply(BigDecimal.valueOf(100)));
        nom.put(key, num(s, "D11", eval).multiply(BigDecimal.valueOf(100)));
    }

    private Path findComisionesFile(LocalDate fechaCorte) {
        String[] contains = {"comisión fpo", "comision fpo", "comisión fpo desde 2003", "comision fpo desde 2003"};
        for (String c : contains) {
            try { return locator.findRequired(c, fechaCorte); } catch (Exception ignored) {}
            try { return locator.findRequired(c); } catch (Exception ignored) {}
        }
        throw new IllegalArgumentException("No se encontró archivo de Comisión FPO desde 2003");
    }

    private Path findInDirContains(Path dir, String contains) throws Exception {
        try (var s = Files.list(dir)) {
            return s.filter(Files::isRegularFile)
                    .filter(p -> normalize(p.getFileName().toString()).contains(normalize(contains)))
                    .findFirst().orElse(null);
        }
    }

    private RowData readBalanceRow(Sheet ws, boolean allowAlt) {
        int cProt = findHeaderCol(ws, "PROTECCION");
        if (cProt < 0) cProt = findHeaderCol(ws, "PROTECCIÓN");
        int cPorv = findHeaderCol(ws, "PORVENIR");
        int cSkan = findHeaderCol(ws, "SKANDIA");
        int cAlt = findHeaderCol(ws, "SKANDIA_ALT");
        int cColf = findHeaderCol(ws, "CITI COLFONDOS");
        if (cColf < 0) cColf = findHeaderCol(ws, "COLFONDOS");
        int cSis = findHeaderCol(ws, "SISTEMA");
        if (cProt < 0 || cPorv < 0 || cSkan < 0 || cColf < 0 || cSis < 0) return new RowData();

        int headerRow = findHeaderRow(ws, cSis);
        int last = ws.getLastRowNum();
        int start = headerRow + 1;
        int end = Math.min(headerRow + 60, last);
        int bestRow = -1;
        BigDecimal max = BigDecimal.valueOf(-1);
        for (int r = start; r <= end; r++) {
            BigDecimal v = num(ws, r + 1, cSis + 1, null);
            if (v.compareTo(max) > 0) { max = v; bestRow = r; }
        }
        if (bestRow < 0) return new RowData();
        RowData d = new RowData();
        d.colf = num(ws, bestRow + 1, cColf + 1, null);
        d.porv = num(ws, bestRow + 1, cPorv + 1, null);
        d.prot = num(ws, bestRow + 1, cProt + 1, null);
        d.skan = num(ws, bestRow + 1, cSkan + 1, null);
        d.alt = (allowAlt && cAlt >= 0) ? num(ws, bestRow + 1, cAlt + 1, null) : BigDecimal.ZERO;
        return d;
    }

    private int findHeaderRow(Sheet ws, int colIdx) {
        for (int r = 0; r <= Math.min(ws.getLastRowNum(), 100); r++) {
            Row row = ws.getRow(r);
            if (row == null) continue;
            Cell c = row.getCell(colIdx);
            if (c == null) continue;
            String t = normalize(new DataFormatter().formatCellValue(c));
            if (t.contains("sistema")) return r;
        }
        return 0;
    }

    private int findHeaderCol(Sheet ws, String text) {
        DataFormatter fmt = new DataFormatter();
        String target = normalize(text);
        for (int r = 0; r <= Math.min(ws.getLastRowNum(), 100); r++) {
            Row row = ws.getRow(r);
            if (row == null) continue;
            for (Cell c : row) {
                String v = normalize(fmt.formatCellValue(c));
                if (v.contains(target)) return c.getColumnIndex();
            }
        }
        return -1;
    }

    private Sheet getSheetIgnoreCase(Workbook wb, String name) {
        for (int i = 0; i < wb.getNumberOfSheets(); i++) {
            Sheet sheet = wb.getSheetAt(i);
            if (sheet.getSheetName().equalsIgnoreCase(name)) return sheet;
        }
        return null;
    }

    private void setDate(Sheet sheet, String ref, LocalDate date) { cell(sheet, ref).setCellValue(java.sql.Date.valueOf(date)); }
    private void setNumeric(Sheet sheet, String ref, double value) { cell(sheet, ref).setCellValue(value); }
    private void setText(Sheet sheet, String ref, String value) { cell(sheet, ref).setCellValue(value); }

    private BigDecimal num(Sheet sheet, String ref, FormulaEvaluator eval) {
        return num(cell(sheet, ref), eval);
    }

    private BigDecimal num(Sheet sheet, int row1, int col1, FormulaEvaluator eval) {
        Row r = sheet.getRow(row1 - 1);
        if (r == null) return BigDecimal.ZERO;
        Cell c = r.getCell(col1 - 1);
        if (c == null) return BigDecimal.ZERO;
        return num(c, eval);
    }

    private BigDecimal num(Cell c, FormulaEvaluator eval) {
        if (eval != null && c.getCellType() == CellType.FORMULA) {
            try {
                CellValue cv = eval.evaluate(c);
                if (cv != null) {
                    return switch (cv.getCellType()) {
                        case NUMERIC -> BigDecimal.valueOf(cv.getNumberValue());
                        case STRING -> parseDecimal(cv.getStringValue());
                        case BOOLEAN -> cv.getBooleanValue() ? BigDecimal.ONE : BigDecimal.ZERO;
                        default -> formulaCachedValue(c);
                    };
                }
            } catch (RuntimeException ex) {
                return formulaCachedValue(c);
            }
            return formulaCachedValue(c);
        }
        return switch (c.getCellType()) {
            case NUMERIC -> BigDecimal.valueOf(c.getNumericCellValue());
            case STRING -> parseDecimal(c.getStringCellValue());
            case BOOLEAN -> c.getBooleanCellValue() ? BigDecimal.ONE : BigDecimal.ZERO;
            default -> BigDecimal.ZERO;
        };
    }

    private BigDecimal formulaCachedValue(Cell c) {
        try {
            return switch (c.getCachedFormulaResultType()) {
                case NUMERIC -> BigDecimal.valueOf(c.getNumericCellValue());
                case STRING -> parseDecimal(c.getStringCellValue());
                case BOOLEAN -> c.getBooleanCellValue() ? BigDecimal.ONE : BigDecimal.ZERO;
                default -> BigDecimal.ZERO;
            };
        } catch (RuntimeException ignored) {
            return BigDecimal.ZERO;
        }
    }

    private BigDecimal parseNumber(Cell cell, DataFormatter formatter) {
        if (cell == null) return BigDecimal.ZERO;
        try { return BigDecimal.valueOf(cell.getNumericCellValue()); }
        catch (Exception ignored) {
            String txt = formatter.formatCellValue(cell).replace(".", "").replace(",", ".").trim();
            if (txt.isBlank()) return BigDecimal.ZERO;
            try { return new BigDecimal(txt); } catch (Exception e) { return BigDecimal.ZERO; }
        }
    }

    private BigDecimal parseDecimal(String s) {
        if (s == null) return BigDecimal.ZERO;
        String n = s.trim().replace(".", "").replace(",", ".");
        if (n.isBlank()) return BigDecimal.ZERO;
        try { return new BigDecimal(n); } catch (Exception e) { return BigDecimal.ZERO; }
    }

    private BigDecimal safeDivide(BigDecimal n, BigDecimal d) {
        if (n == null || d == null || d.signum() == 0) return BigDecimal.ZERO;
        return n.divide(d, 8, java.math.RoundingMode.HALF_UP);
    }

    private Cell cell(Sheet sheet, String ref) {
        CellReference cr = new CellReference(ref);
        Row row = sheet.getRow(cr.getRow());
        if (row == null) row = sheet.createRow(cr.getRow());
        Cell cell = row.getCell(cr.getCol());
        if (cell == null) cell = row.createCell(cr.getCol());
        return cell;
    }

    private String normalize(String value) {
        if (value == null) return "";
        String n = Normalizer.normalize(value, Normalizer.Form.NFD).replaceAll("\\p{M}", "");
        return n.toLowerCase(Locale.ROOT).trim();
    }

    private static class RowData {
        BigDecimal colf = BigDecimal.ZERO;
        BigDecimal porv = BigDecimal.ZERO;
        BigDecimal prot = BigDecimal.ZERO;
        BigDecimal skan = BigDecimal.ZERO;
        BigDecimal alt = BigDecimal.ZERO;
    }
}
