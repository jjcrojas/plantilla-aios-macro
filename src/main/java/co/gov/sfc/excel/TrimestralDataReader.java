package co.gov.sfc.excel;

import co.gov.sfc.insumos.InsumosLocator;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
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
import java.util.HashMap;
import java.util.Locale;
import java.util.Map;

@Component
public class TrimestralDataReader {

    private static final Logger log = LoggerFactory.getLogger(TrimestralDataReader.class);

    private final MensualDataReader mensualDataReader;
    private final InsumosLocator locator;

    public TrimestralDataReader(MensualDataReader mensualDataReader, InsumosLocator locator) {
        this.mensualDataReader = mensualDataReader;
        this.locator = locator;
    }

    public TrimestralData read(LocalDate fechaCorte) {
        MensualData mensual = mensualDataReader.read(fechaCorte);
        Map<String, BigDecimal> cotizantes = readCotizantesFrom491(fechaCorte);
        Map<String, BigDecimal> traspasos = readTraspasosFrom493(fechaCorte);

        String etiquetaFecha = fechaCorte.getMonth().getDisplayName(TextStyle.SHORT, new Locale("es", "CO"))
                .replace(".", "")
                .toLowerCase() + "-" + String.format("%02d", fechaCorte.getYear() % 100);

        return new TrimestralData(
                etiquetaFecha,
                cotizantes.getOrDefault("colfondos", BigDecimal.ZERO),
                cotizantes.getOrDefault("porvenir", BigDecimal.ZERO),
                cotizantes.getOrDefault("proteccion", BigDecimal.ZERO),
                cotizantes.getOrDefault("skandia", BigDecimal.ZERO),
                mensual.vrFondo().max(BigDecimal.ZERO),
                traspasos.getOrDefault("colfondos", BigDecimal.ZERO),
                traspasos.getOrDefault("porvenir", BigDecimal.ZERO),
                traspasos.getOrDefault("proteccion", BigDecimal.ZERO),
                traspasos.getOrDefault("skandia", BigDecimal.ZERO),
                mensual.tmpNominal1().multiply(BigDecimal.valueOf(100)),
                mensual.tmpReal1().multiply(BigDecimal.valueOf(100))
        );
    }

    private Map<String, BigDecimal> readTraspasosFrom493(LocalDate fechaCorte) {
        Path file493 = locator.findRequired("493", fechaCorte);
        Map<String, BigDecimal> out = new HashMap<>();

        try (Workbook wb = WorkbookFactory.create(file493.toFile(), null, true)) {
            Sheet sheet = getSheetIgnoreCase(wb, "Traslados Entre AFP");
            if (sheet == null) {
                throw new IllegalStateException("No existe hoja 'Traslados Entre AFP' en Formato 493");
            }
            FormulaEvaluator evaluator = wb.getCreationHelper().createFormulaEvaluator();
            setDate(sheet, "B11", fechaCorte);

            out.put("colfondos", readTraspasosByCode(sheet, evaluator, 10));
            out.put("proteccion", readTraspasosByCode(sheet, evaluator, 2));
            out.put("porvenir", readTraspasosByCode(sheet, evaluator, 3));
            out.put("skandia", readTraspasosByCode(sheet, evaluator, 9));

            return out;
        } catch (Exception e) {
            throw new IllegalStateException("Error leyendo traspasos por AFP desde Formato 493", e);
        }
    }

    private BigDecimal readTraspasosByCode(Sheet sheet, FormulaEvaluator evaluator, int afpCode) {
        // Macro: D4=codAFP y leer BQ11. Aquí forzamos evaluación de los componentes
        // (M11, AA11, AO11, BC11) para evitar valores cacheados de BQ11.
        setNumeric(sheet, "D4", afpCode);
        evaluator.clearAllCachedResultValues();

        BigDecimal total = num(sheet, "M11", evaluator)
                .add(num(sheet, "AA11", evaluator))
                .add(num(sheet, "AO11", evaluator))
                .add(num(sheet, "BC11", evaluator));

        if (total.signum() != 0) {
            return total;
        }

        // Fallback: algunas plantillas pueden almacenar D4 como texto.
        setText(sheet, "D4", String.valueOf(afpCode));
        evaluator.clearAllCachedResultValues();
        total = num(sheet, "M11", evaluator)
                .add(num(sheet, "AA11", evaluator))
                .add(num(sheet, "AO11", evaluator))
                .add(num(sheet, "BC11", evaluator));

        if (total.signum() != 0) {
            return total;
        }

        // Último fallback: valor directo en BQ11.
        return num(sheet, "BQ11", evaluator);
    }

    private Map<String, BigDecimal> readCotizantesFrom491(LocalDate fechaCorte) {
        Path local491 = Path.of("insumos_ejemplo", "Serie_Formato_ 491 AFILIADOS AFP.xlsm");
        if (!Files.exists(local491) || !Files.isRegularFile(local491)) {
            throw new IllegalStateException("No se encontró Formato 491 en ./insumos_ejemplo/Serie_Formato_ 491 AFILIADOS AFP.xlsm");
        }

        Map<String, BigDecimal> out = new HashMap<>();
        try (Workbook wb = WorkbookFactory.create(local491.toFile(), null, true)) {
            Sheet sheet = wb.getSheet("multifondos");
            if (sheet == null) {
                throw new IllegalStateException("No existe hoja 'multifondos' en Formato 491");
            }

            int entidadCol = -1;
            int cotizantesCol = -1;
            int headerRow = -1;
            DataFormatter formatter = new DataFormatter();

            for (int r = 0; r <= Math.min(sheet.getLastRowNum(), 200); r++) {
                Row row = sheet.getRow(r);
                if (row == null) continue;
                for (Cell cell : row) {
                    String txt = normalize(formatter.formatCellValue(cell));
                    if (txt.contains("entidad")) entidadCol = cell.getColumnIndex();
                    if (txt.contains("cotizantes")) cotizantesCol = cell.getColumnIndex();
                }
                if (entidadCol >= 0 && cotizantesCol >= 0) {
                    headerRow = r;
                    break;
                }
            }

            if (headerRow < 0) {
                log.warn("No se ubicaron encabezados ENTIDAD/COTIZANTES en 491 para fechaCorte={}", fechaCorte);
                return out;
            }

            for (int r = headerRow + 1; r <= Math.min(sheet.getLastRowNum(), headerRow + 150); r++) {
                Row row = sheet.getRow(r);
                if (row == null) continue;
                String entidad = normalize(formatter.formatCellValue(row.getCell(entidadCol)));
                if (entidad.isBlank()) continue;
                BigDecimal cot = parseNumber(row.getCell(cotizantesCol), formatter);
                if (entidad.contains("colfond")) out.put("colfondos", cot);
                else if (entidad.contains("porvenir")) out.put("porvenir", cot);
                else if (entidad.contains("protec")) out.put("proteccion", cot);
                else if (entidad.contains("skand")) out.put("skandia", cot);
            }

            return out;
        } catch (Exception e) {
            throw new IllegalStateException("Error leyendo cotizantes por AFP desde Formato 491", e);
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

    private void setDate(Sheet sheet, String ref, LocalDate date) {
        cell(sheet, ref).setCellValue(java.sql.Date.valueOf(date));
    }

    private void setNumeric(Sheet sheet, String ref, double value) {
        cell(sheet, ref).setCellValue(value);
    }

    private void setText(Sheet sheet, String ref, String value) {
        cell(sheet, ref).setCellValue(value);
    }

    private BigDecimal num(Sheet sheet, String ref, FormulaEvaluator eval) {
        Cell c = cell(sheet, ref);
        if (eval != null && c.getCellType() == CellType.FORMULA) {
            var cv = eval.evaluate(c);
            if (cv == null) return BigDecimal.ZERO;
            return switch (cv.getCellType()) {
                case NUMERIC -> BigDecimal.valueOf(cv.getNumberValue());
                case STRING -> parseDecimal(cv.getStringValue());
                case BOOLEAN -> cv.getBooleanValue() ? BigDecimal.ONE : BigDecimal.ZERO;
                default -> BigDecimal.ZERO;
            };
        }
        return switch (c.getCellType()) {
            case NUMERIC -> BigDecimal.valueOf(c.getNumericCellValue());
            case STRING -> parseDecimal(c.getStringCellValue());
            case BOOLEAN -> c.getBooleanCellValue() ? BigDecimal.ONE : BigDecimal.ZERO;
            default -> BigDecimal.ZERO;
        };
    }

    private BigDecimal parseDecimal(String s) {
        if (s == null) return BigDecimal.ZERO;
        String n = s.trim().replace(".", "").replace(",", ".");
        if (n.isBlank()) return BigDecimal.ZERO;
        try {
            return new BigDecimal(n);
        } catch (Exception e) {
            return BigDecimal.ZERO;
        }
    }

    private Cell cell(Sheet sheet, String ref) {
        CellReference cr = new CellReference(ref);
        Row row = sheet.getRow(cr.getRow());
        if (row == null) row = sheet.createRow(cr.getRow());
        Cell cell = row.getCell(cr.getCol());
        if (cell == null) cell = row.createCell(cr.getCol());
        return cell;
    }

    private BigDecimal parseNumber(Cell cell, DataFormatter formatter) {
        if (cell == null) return BigDecimal.ZERO;
        try {
            return BigDecimal.valueOf(cell.getNumericCellValue());
        } catch (Exception ignored) {
            String txt = formatter.formatCellValue(cell).replace(".", "").replace(",", ".").trim();
            if (txt.isBlank()) return BigDecimal.ZERO;
            try {
                return new BigDecimal(txt);
            } catch (Exception e) {
                return BigDecimal.ZERO;
            }
        }
    }

    private String normalize(String value) {
        if (value == null) return "";
        String n = Normalizer.normalize(value, Normalizer.Form.NFD).replaceAll("\\p{M}", "");
        return n.toLowerCase(Locale.ROOT).trim();
    }
}
