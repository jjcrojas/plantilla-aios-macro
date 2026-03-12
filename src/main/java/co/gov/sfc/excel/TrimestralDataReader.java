package co.gov.sfc.excel;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
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

    public TrimestralDataReader(MensualDataReader mensualDataReader) {
        this.mensualDataReader = mensualDataReader;
    }

    public TrimestralData read(LocalDate fechaCorte) {
        MensualData mensual = mensualDataReader.read(fechaCorte);
        Map<String, BigDecimal> cotizantes = readCotizantesFrom491(fechaCorte);

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
                mensual.traspasosSistema().max(BigDecimal.ZERO),
                mensual.tmpNominal1().multiply(BigDecimal.valueOf(100)),
                mensual.tmpReal1().multiply(BigDecimal.valueOf(100))
        );
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

