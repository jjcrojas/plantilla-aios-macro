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

@Component
public class MensualDataReader {

    private static final Logger log = LoggerFactory.getLogger(MensualDataReader.class);
    private final InsumosLocator locator;

    public MensualDataReader(InsumosLocator locator, AiosProperties properties) {
        this.locator = locator;
        // Evitar asignaciones gigantes en POI que pueden terminar en OOM con archivos grandes.
        // 100 MB es suficiente para los insumos actuales y más conservador en memoria.
        IOUtils.setByteArrayMaxOverride(100_000_000);
    }

    public MensualData read(LocalDate fechaCorte) {
        log.info("Iniciando lectura de insumos para fechaCorte={}", fechaCorte);

        BigDecimal hombres;
        BigDecimal mujeres;
        BigDecimal aportantes;
        BigDecimal consFdosAdmon;

        var file491 = locator.findRequired("491", fechaCorte);
        try (Workbook wb = WorkbookFactory.create(file491.toFile(), null, true)) {
            FormulaEvaluator evaluator = wb.getCreationHelper().createFormulaEvaluator();
            Sheet informe = wb.getSheet("informe de prensa");
            Sheet multifondos = wb.getSheet("multifondos");

            setDate(informe, "C3", fechaCorte);
            setDate(multifondos, "C4", fechaCorte);
            log.info("Formato 491 actualizado con fechaCorte={} en informe!C3 y multifondos!C4", fechaCorte);

            evaluator.clearAllCachedResultValues();

            // Igual que macro VBA: afiliados = hombres + mujeres (informe de prensa).
            // Para evitar quedarse con valores cacheados del archivo guardado (ej. diciembre),
            // se recalcula desde filas de detalle dependientes de C3 (C7:C10 y D7:D10).
            hombres = sumRange(informe, evaluator, 7, 10, 3);
            mujeres = sumRange(informe, evaluator, 7, 10, 4);
            log.info("Afiliados desde 491 para fechaCorte={}: hombres={}, mujeres={}, total={}", fechaCorte, hombres, mujeres, hombres.add(mujeres));

            aportantes = num(multifondos, "E25", evaluator);
            log.info("Aportantes desde 491 para fechaCorte={}: {}", fechaCorte, aportantes);
            var j8 = num(multifondos, "J8", evaluator);
            var j9 = num(multifondos, "J9", evaluator);
            var j12 = num(multifondos, "J12", evaluator);
            consFdosAdmon = j12.signum() == 0 ? BigDecimal.ZERO : j8.add(j9).divide(j12, 8, RoundingMode.HALF_UP).multiply(BigDecimal.valueOf(100));
        } catch (Exception e) {
            throw new IllegalStateException("Error leyendo Formato 491", e);
        }
        log.info("Lectura Formato 491 completada para fechaCorte={}", fechaCorte);

        BigDecimal traspasosSistema = BigDecimal.ZERO;
        var file493 = locator.findRequired("493", fechaCorte);
        try (Workbook wb = WorkbookFactory.create(file493.toFile(), null, true)) {
            FormulaEvaluator evaluator = wb.getCreationHelper().createFormulaEvaluator();
            Sheet tras = wb.getSheet("Traslados Entre AFP");
            if (tras == null) {
                throw new IllegalStateException("No existe hoja 'Traslados Entre AFP' en Formato 493");
            }
            setDate(tras, "B11", fechaCorte);
            setNumeric(tras, "D4", 99);
            traspasosSistema = num(tras, "BQ11", evaluator);
        } catch (Exception e) {
            log.warn("No fue posible leer Formato 493; se usará 0 en traspasos_sistema. Causa: {}", e.getMessage());
        }
        log.info("Lectura Formato 493 completada para fechaCorte={}", fechaCorte);

        BigDecimal tmpReal1;
        BigDecimal tmpNominal1;
        var rentFile = locator.findRequired("Rent_Vr_Uni_Moderado", fechaCorte);
        try (Workbook wb = WorkbookFactory.create(rentFile.toFile(), null, true)) {
            Sheet consolidado = getSheetIgnoreCase(wb, "Consolidado");
            if (consolidado == null) consolidado = wb.getSheetAt(0);

            LocalDate fechaInicial = fechaCorte.minusYears(1);
            // Igual que macro: D5 = fecha final, D4 = fecha inicial.
            setDate(consolidado, "D5", fechaCorte);
            setDate(consolidado, "D4", fechaInicial);

            tmpNominal1 = readRentabilidadNominal(consolidado, fechaInicial, fechaCorte);
            tmpReal1 = readRentabilidadReal(consolidado, fechaCorte);
            log.info("Rentabilidad moderado con fechaCorte={}: consolidado!D4={}, D5={}, D11(nominal)={}, D10(real)={}",
                    fechaCorte, fechaInicial, fechaCorte, tmpNominal1, tmpReal1);
        } catch (Exception e) {
            throw new IllegalStateException("Error leyendo rentabilidad moderado", e);
        }
        log.info("Lectura rentabilidad completada para fechaCorte={}", fechaCorte);

        BigDecimal vrFondo = BigDecimal.ZERO;
        BigDecimal porcVrFondo = BigDecimal.ZERO;
        var sistemaTotal = locator.findRequired("SISTEMA TOTAL", fechaCorte);
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
