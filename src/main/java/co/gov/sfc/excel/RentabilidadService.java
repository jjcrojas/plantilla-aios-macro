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
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.LocalDate;
import java.time.YearMonth;
import java.time.ZoneId;
import java.time.temporal.ChronoUnit;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Comparator;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.NavigableMap;
import java.util.TreeMap;
import java.util.stream.Collectors;
import java.util.stream.Stream;

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
        NavigableMap<LocalDate, BigDecimal> nav = readNavPromedio(valoresFondoModerFile, fechaInicio, fechaCorte);
        NavigableMap<YearMonth, BigDecimal> ipc = readIpcSeries(rentModeradoFile);
        return calcular(fechaInicio, fechaCorte, nav, ipc);
    }

    private RentabilidadResultado calcular(
            LocalDate fechaInicio,
            LocalDate fechaFin,
            NavigableMap<LocalDate, BigDecimal> nav,
            NavigableMap<YearMonth, BigDecimal> ipc
    ) {
        var navIniEntry = nearestStartValue(nav, fechaInicio);
        var navFinEntry = nav.floorEntry(fechaFin);
        if (navIniEntry == null || navFinEntry == null) {
            throw new IllegalStateException("No hay NAV suficiente para calcular horizonte. "
                    + "inicio=" + fechaInicio + " fin=" + fechaFin
                    + " navIni=" + (navIniEntry == null ? "null" : navIniEntry.getKey())
                    + " navFin=" + (navFinEntry == null ? "null" : navFinEntry.getKey()));
        }
        BigDecimal navIni = navIniEntry.getValue();
        BigDecimal navFin = navFinEntry.getValue();

        long dias = Math.max(1, ChronoUnit.DAYS.between(fechaInicio, fechaFin));
        double navFactor = navFin.divide(navIni, 16, RoundingMode.HALF_UP).doubleValue();
        BigDecimal nominal = BigDecimal.valueOf(Math.pow(navFactor, 365d / (double) dias) - 1d);

        var ipcIniEntry = ipc.floorEntry(YearMonth.from(fechaInicio));
        var ipcFinEntry = ipc.floorEntry(YearMonth.from(fechaFin));
        if (ipcIniEntry == null || ipcFinEntry == null) {
            throw new IllegalStateException("No hay IPC suficiente para calcular horizonte. "
                    + "inicio=" + fechaInicio + " fin=" + fechaFin
                    + " ipcIni=" + (ipcIniEntry == null ? "null" : ipcIniEntry.getKey())
                    + " ipcFin=" + (ipcFinEntry == null ? "null" : ipcFinEntry.getKey()));
        }
        BigDecimal ipcIni = ipcIniEntry.getValue();
        BigDecimal ipcFin = ipcFinEntry.getValue();
        validarCrecimientoIpc(ipcIni, ipcFin, fechaInicio, fechaFin);
        BigDecimal inflacion = ipcFin.divide(ipcIni, 16, RoundingMode.HALF_UP).subtract(BigDecimal.ONE);
        BigDecimal real = nominal.add(BigDecimal.ONE)
                .divide(inflacion.add(BigDecimal.ONE), 16, RoundingMode.HALF_UP)
                .subtract(BigDecimal.ONE);

        log.info("Rentabilidad NAV calculada: ini={} fin={} navIniDate={} navIni={} navFinDate={} navFin={} ipcIniDate={} ipcIni={} ipcFinDate={} ipcFin={} nominal={} inflacion={} real={}",
                fechaInicio, fechaFin, navIniEntry.getKey(), navIni, navFinEntry.getKey(), navFin,
                ipcIniEntry.getKey(), ipcIni, ipcFinEntry.getKey(), ipcFin, nominal, inflacion, real);
        return new RentabilidadResultado(fechaInicio, fechaFin, nominal, real);
    }

    private NavigableMap<LocalDate, BigDecimal> readNavPromedio(Path file, LocalDate fechaInicio, LocalDate fechaFin) {
            Map<LocalDate, List<BigDecimal>> porFecha = new TreeMap<>();
        List<Path> navFiles = findNavHistoryFiles(file);
        int fondosEsperados = 0;
        try {
            for (Path navFile : navFiles) {
                try (Workbook wb = WorkbookFactory.create(navFile.toFile(), null, true)) {
                    List<String> hojasNav = detectHojasNav(wb);
                    fondosEsperados = Math.max(fondosEsperados, hojasNav.size());
                    for (String nombre : hojasNav) {
                        Sheet s = wb.getSheet(nombre);
                        if (s == null) continue;
                        int last = s.getLastRowNum() + 1;
                        for (int r = 2; r <= last; r++) {
                            Row row = s.getRow(r - 1);
                            if (row == null) continue;
                            LocalDate fecha = cellAsDate(row.getCell(0));
                            BigDecimal nav = cellAsNumber(row.getCell(14)); // columna O
                            if (fecha == null || nav.signum() <= 0) continue;
                            if (fecha.isAfter(fechaFin)) continue;
                            porFecha.computeIfAbsent(fecha, k -> new ArrayList<>()).add(nav);
                        }
                    }
                }
            }
        } catch (Exception e) {
            throw new IllegalStateException("No fue posible leer NAV histórico desde " + file.toAbsolutePath() + ": " + e.getMessage(), e);
        }
        try {
            NavigableMap<LocalDate, BigDecimal> serie = new TreeMap<>();
            int coberturaParcial = 0;
            for (var e : porFecha.entrySet()) {
                BigDecimal sum = e.getValue().stream().reduce(BigDecimal.ZERO, BigDecimal::add);
                BigDecimal avg = sum.divide(BigDecimal.valueOf(e.getValue().size()), 16, RoundingMode.HALF_UP);
                serie.put(e.getKey(), avg);
                if (e.getValue().size() < fondosEsperados) {
                    coberturaParcial++;
                }
            }
            log.info("Serie NAV promedio cargada: file={} fechas={} desde={} hasta={} fondosEsperados={} fechasCoberturaParcial={}",
                    file.toAbsolutePath(), serie.size(),
                    serie.isEmpty() ? null : serie.firstKey(),
                    serie.isEmpty() ? null : serie.lastKey(),
                    fondosEsperados,
                    coberturaParcial);
            if (serie.isEmpty()) {
                throw new IllegalStateException("No hay NAV en el rango solicitado [" + fechaInicio + ", " + fechaFin + "] en " + file.toAbsolutePath());
            }
            return serie;
        } catch (IllegalStateException e) {
            throw e;
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

    private NavigableMap<YearMonth, BigDecimal> readIpcSeries(Path rentModeradoFile) {
        try (Workbook wb = WorkbookFactory.create(rentModeradoFile.toFile(), null, true)) {
            Sheet ipcBr = getSheetIgnoreCase(wb, "IPC_BR");
            if (ipcBr != null) {
                NavigableMap<YearMonth, BigDecimal> serie = readDateValueSheetByMonth(ipcBr, 1, 2);
                if (!serie.isEmpty()) {
                    log.info("Serie IPC_BR cargada: file={} fechas={}", rentModeradoFile.toAbsolutePath(), serie.size());
                    return serie;
                }
            }
            Sheet ipc = getSheetIgnoreCase(wb, "IPC");
            if (ipc != null) {
                NavigableMap<YearMonth, BigDecimal> tasas = readDateValueSheetByMonth(ipc, 1, 2);
                if (!tasas.isEmpty()) {
                    if (isIndexSeries(tasas)) {
                        log.info("Serie IPC cargada como índice directo: file={} fechas={}", rentModeradoFile.toAbsolutePath(), tasas.size());
                        return tasas;
                    }
                    BigDecimal indice = BigDecimal.valueOf(100);
                    NavigableMap<YearMonth, BigDecimal> indices = new TreeMap<>();
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

    private Map.Entry<LocalDate, BigDecimal> nearestStartValue(NavigableMap<LocalDate, BigDecimal> serie, LocalDate fechaInicio) {
        if (serie == null || serie.isEmpty()) return null;
        return serie.floorEntry(fechaInicio);
    }

    private boolean isIndexSeries(NavigableMap<YearMonth, BigDecimal> serie) {
        if (serie.isEmpty()) return false;
        List<BigDecimal> sample = new ArrayList<>(serie.values());
        Collections.sort(sample);
        BigDecimal median = sample.get(sample.size() / 2);
        // Regla simple: un IPC índice suele estar muy por encima de 1; una tasa mensual suele estar < 1.
        return median.compareTo(BigDecimal.valueOf(2)) > 0;
    }

    private NavigableMap<YearMonth, BigDecimal> readDateValueSheetByMonth(Sheet sheet, int dateCol1Based, int valueCol1Based) {
        NavigableMap<YearMonth, BigDecimal> data = new TreeMap<>(Comparator.naturalOrder());
        int last = sheet.getLastRowNum() + 1;
        for (int r = 2; r <= last; r++) {
            Row row = sheet.getRow(r - 1);
            if (row == null) continue;
            LocalDate fecha = cellAsDate(row.getCell(dateCol1Based - 1));
            BigDecimal valor = cellAsNumber(row.getCell(valueCol1Based - 1));
            if (fecha == null || valor.signum() <= 0) continue;
            data.put(YearMonth.from(fecha), valor);
        }
        return data;
    }

    private List<Path> findNavHistoryFiles(Path oneFile) {
        Path historico = findAncestor(oneFile, "Historico_Rent_minima");
        if (historico == null || !Files.isDirectory(historico)) {
            return List.of(oneFile);
        }
        try (Stream<Path> stream = Files.walk(historico, 4)) {
            List<Path> files = stream
                    .filter(Files::isRegularFile)
                    .filter(p -> p.getFileName().toString().toLowerCase(Locale.ROOT).contains("valores_fondo_moder"))
                    .sorted()
                    .collect(Collectors.toList());
            return files.isEmpty() ? List.of(oneFile) : files;
        } catch (Exception e) {
            return List.of(oneFile);
        }
    }

    private Path findAncestor(Path start, String folderName) {
        Path p = start;
        while (p != null) {
            Path name = p.getFileName();
            if (name != null && name.toString().equalsIgnoreCase(folderName)) {
                return p;
            }
            p = p.getParent();
        }
        return null;
    }

    private void validarCrecimientoIpc(BigDecimal ipcIni, BigDecimal ipcFin, LocalDate fechaInicio, LocalDate fechaFin) {
        if (ipcIni.signum() <= 0 || ipcFin.signum() <= 0) {
            throw new IllegalStateException("IPC inválido (<=0): ini=" + ipcIni + " fin=" + ipcFin);
        }
        if (ipcFin.compareTo(ipcIni.multiply(BigDecimal.valueOf(5))) > 0) {
            throw new IllegalStateException("IPC inválido: crecimiento irreal " + ipcIni + " -> " + ipcFin
                    + " para rango " + fechaInicio + " a " + fechaFin);
        }
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
