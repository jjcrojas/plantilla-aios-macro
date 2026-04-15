package co.gov.sfc.excel;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Service;

import java.math.BigDecimal;
import java.math.RoundingMode;
import java.nio.file.Path;
import java.time.LocalDate;
import java.time.ZoneId;
import java.time.temporal.ChronoUnit;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.NavigableMap;
import java.util.TreeMap;
import java.util.stream.Collectors;

@Service
public class RentabilidadService {

    private static final Logger log = LoggerFactory.getLogger(RentabilidadService.class);

    private static final List<String> HOJAS_NAV = List.of(
            "CO_Vr_Uni", "OM_Vr_Uni", "PRO_Vr_Uni", "PO_Vr_Uni",
            "Colfondos", "OldMutual", "Protección", "Porvenir"
    );

    public RentabilidadResultado calcularRentabilidad(
            Path valoresFondoModerFile,
            Path rentModeradoFile,
            LocalDate fechaCorte,
            int horizonteAnios
    ) {
        LocalDate fechaInicio = fechaCorte.minusYears(horizonteAnios);
        NavigableMap<LocalDate, BigDecimal> nav = readNavPromedio(valoresFondoModerFile);
        NavigableMap<LocalDate, BigDecimal> ipc = readIpcSeries(rentModeradoFile);
        return calcular(fechaInicio, fechaCorte, nav, ipc);
    }

    private RentabilidadResultado calcular(
            LocalDate fechaInicio,
            LocalDate fechaFin,
            NavigableMap<LocalDate, BigDecimal> nav,
            NavigableMap<LocalDate, BigDecimal> ipc
    ) {
        BigDecimal navIni = floorValue(nav, fechaInicio);
        BigDecimal navFin = floorValue(nav, fechaFin);
        if (navIni.signum() <= 0 || navFin.signum() <= 0) {
            log.warn("Rentabilidad NAV sin datos válidos: ini={} fin={} navIni={} navFin={}",
                    fechaInicio, fechaFin, navIni, navFin);
            return new RentabilidadResultado(fechaInicio, fechaFin, BigDecimal.ZERO, BigDecimal.ZERO);
        }

        long dias = Math.max(1, ChronoUnit.DAYS.between(fechaInicio, fechaFin));
        double navFactor = navFin.divide(navIni, 16, RoundingMode.HALF_UP).doubleValue();
        BigDecimal nominal = BigDecimal.valueOf(Math.pow(navFactor, 365d / (double) dias) - 1d);

        BigDecimal ipcIni = floorValue(ipc, fechaInicio);
        BigDecimal ipcFin = floorValue(ipc, fechaFin);
        BigDecimal real;
        if (ipcIni.signum() > 0 && ipcFin.signum() > 0) {
            double ipcFactor = ipcFin.divide(ipcIni, 16, RoundingMode.HALF_UP).doubleValue();
            double realFactor = navFactor / ipcFactor;
            real = BigDecimal.valueOf(Math.pow(realFactor, 365d / (double) dias) - 1d);
        } else {
            real = nominal;
            log.warn("Rentabilidad real sin IPC válido para ini={} fin={}; se usa nominal como fallback.", fechaInicio, fechaFin);
        }

        log.info("Rentabilidad NAV calculada: ini={} fin={} navIni={} navFin={} ipcIni={} ipcFin={} nominal={} real={}",
                fechaInicio, fechaFin, navIni, navFin, ipcIni, ipcFin, nominal, real);
        return new RentabilidadResultado(fechaInicio, fechaFin, nominal, real);
    }

    private NavigableMap<LocalDate, BigDecimal> readNavPromedio(Path file) {
        try (Workbook wb = WorkbookFactory.create(file.toFile(), null, true)) {
            Map<LocalDate, List<BigDecimal>> porFecha = new TreeMap<>();
            for (String nombre : detectHojasNav(wb)) {
                Sheet s = wb.getSheet(nombre);
                if (s == null) continue;
                int last = s.getLastRowNum() + 1;
                for (int r = 2; r <= last; r++) {
                    Row row = s.getRow(r - 1);
                    if (row == null) continue;
                    LocalDate fecha = cellAsDate(row.getCell(0));
                    BigDecimal nav = cellAsNumber(row.getCell(14)); // columna O
                    if (fecha == null || nav.signum() <= 0) continue;
                    porFecha.computeIfAbsent(fecha, k -> new ArrayList<>()).add(nav);
                }
            }
            NavigableMap<LocalDate, BigDecimal> serie = new TreeMap<>();
            for (var e : porFecha.entrySet()) {
                BigDecimal sum = e.getValue().stream().reduce(BigDecimal.ZERO, BigDecimal::add);
                BigDecimal avg = sum.divide(BigDecimal.valueOf(e.getValue().size()), 16, RoundingMode.HALF_UP);
                serie.put(e.getKey(), avg);
            }
            log.info("Serie NAV promedio cargada: file={} fechas={} desde={} hasta={}",
                    file.toAbsolutePath(), serie.size(),
                    serie.isEmpty() ? null : serie.firstKey(),
                    serie.isEmpty() ? null : serie.lastKey());
            return serie;
        } catch (Exception e) {
            log.warn("No fue posible leer NAV desde {}: {}", file.toAbsolutePath(), e.getMessage());
            return new TreeMap<>();
        }
    }

    private List<String> detectHojasNav(Workbook wb) {
        List<String> exactas = HOJAS_NAV.stream().filter(n -> wb.getSheet(n) != null).toList();
        if (!exactas.isEmpty()) return exactas;
        List<String> porPatron = new ArrayList<>();
        for (int i = 0; i < wb.getNumberOfSheets(); i++) {
            String name = wb.getSheetAt(i).getSheetName();
            String n = name.toLowerCase(Locale.ROOT);
            if (n.contains("vr_uni") || n.contains("colfondos") || n.contains("oldmutual")
                    || n.contains("prote") || n.contains("porvenir")) {
                porPatron.add(name);
            }
        }
        if (!porPatron.isEmpty()) return porPatron;
        return List.of(wb.getSheetAt(0).getSheetName());
    }

    private NavigableMap<LocalDate, BigDecimal> readIpcSeries(Path rentModeradoFile) {
        try (Workbook wb = WorkbookFactory.create(rentModeradoFile.toFile(), null, true)) {
            Sheet ipcBr = getSheetIgnoreCase(wb, "IPC_BR");
            if (ipcBr != null) {
                NavigableMap<LocalDate, BigDecimal> serie = readDateValueSheet(ipcBr, 1, 2);
                if (!serie.isEmpty()) {
                    log.info("Serie IPC_BR cargada: file={} fechas={}", rentModeradoFile.toAbsolutePath(), serie.size());
                    return serie;
                }
            }
            Sheet ipc = getSheetIgnoreCase(wb, "IPC");
            if (ipc != null) {
                NavigableMap<LocalDate, BigDecimal> tasas = readDateValueSheet(ipc, 1, 2);
                if (!tasas.isEmpty()) {
                    BigDecimal indice = BigDecimal.valueOf(100);
                    NavigableMap<LocalDate, BigDecimal> indices = new TreeMap<>();
                    for (var e : tasas.entrySet().stream().sorted(Map.Entry.comparingByKey()).collect(Collectors.toList())) {
                        indice = indice.multiply(BigDecimal.ONE.add(e.getValue())).setScale(16, RoundingMode.HALF_UP);
                        indices.put(e.getKey(), indice);
                    }
                    log.info("Serie IPC (tasas->índice) cargada: file={} fechas={}", rentModeradoFile.toAbsolutePath(), indices.size());
                    return indices;
                }
            }
        } catch (Exception e) {
            log.warn("No fue posible leer IPC desde {}: {}", rentModeradoFile.toAbsolutePath(), e.getMessage());
        }
        return new TreeMap<>();
    }

    private NavigableMap<LocalDate, BigDecimal> readDateValueSheet(Sheet sheet, int dateCol1Based, int valueCol1Based) {
        NavigableMap<LocalDate, BigDecimal> data = new TreeMap<>(Comparator.naturalOrder());
        int last = sheet.getLastRowNum() + 1;
        for (int r = 2; r <= last; r++) {
            Row row = sheet.getRow(r - 1);
            if (row == null) continue;
            LocalDate fecha = cellAsDate(row.getCell(dateCol1Based - 1));
            BigDecimal valor = cellAsNumber(row.getCell(valueCol1Based - 1));
            if (fecha == null || valor.signum() <= 0) continue;
            data.put(fecha, valor);
        }
        return data;
    }

    private BigDecimal floorValue(NavigableMap<LocalDate, BigDecimal> serie, LocalDate fecha) {
        if (serie == null || serie.isEmpty()) return BigDecimal.ZERO;
        var exacta = serie.get(fecha);
        if (exacta != null) return exacta;
        var floor = serie.floorEntry(fecha);
        return floor == null ? BigDecimal.ZERO : floor.getValue();
    }

    private LocalDate cellAsDate(Cell c) {
        if (c == null) return null;
        try {
            return switch (c.getCellType()) {
                case NUMERIC -> c.getDateCellValue().toInstant().atZone(ZoneId.systemDefault()).toLocalDate();
                case STRING -> {
                    String s = c.getStringCellValue();
                    if (s == null || s.isBlank()) yield null;
                    yield LocalDate.parse(s.trim());
                }
                case FORMULA -> {
                    if (c.getCachedFormulaResultType() == org.apache.poi.ss.usermodel.CellType.NUMERIC) {
                        yield c.getDateCellValue().toInstant().atZone(ZoneId.systemDefault()).toLocalDate();
                    }
                    yield null;
                }
                default -> null;
            };
        } catch (Exception ignore) {
            return null;
        }
    }

    private BigDecimal cellAsNumber(Cell c) {
        if (c == null) return BigDecimal.ZERO;
        try {
            return switch (c.getCellType()) {
                case NUMERIC -> BigDecimal.valueOf(c.getNumericCellValue());
                case FORMULA -> c.getCachedFormulaResultType() == org.apache.poi.ss.usermodel.CellType.NUMERIC
                        ? BigDecimal.valueOf(c.getNumericCellValue())
                        : BigDecimal.ZERO;
                case STRING -> parseDecimal(c.getStringCellValue());
                default -> BigDecimal.ZERO;
            };
        } catch (Exception ignore) {
            return BigDecimal.ZERO;
        }
    }

    private BigDecimal parseDecimal(String value) {
        if (value == null) return BigDecimal.ZERO;
        String v = value.trim();
        if (v.isEmpty()) return BigDecimal.ZERO;
        try {
            return new BigDecimal(v.replace(",", ""));
        } catch (Exception e) {
            return BigDecimal.ZERO;
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

    public record RentabilidadResultado(
            LocalDate fechaInicio,
            LocalDate fechaFin,
            BigDecimal rentabilidadNominal,
            BigDecimal rentabilidadReal
    ) {}
}
