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
        NavSeries navSeries = readNavPromedio(valoresFondoModerFile, fechaInicio, fechaCorte);
        if (navSeries.values().floorEntry(fechaInicio) == null) {
            NavSeries navConsolidado = readNavFromRentConsolidado(rentModeradoFile, fechaCorte);
            if (!navConsolidado.values().isEmpty()) {
                navSeries.values().putAll(navConsolidado.values());
                navSeries.sourceByDate().putAll(navConsolidado.sourceByDate());
                navSeries.contributorsByDate().putAll(navConsolidado.contributorsByDate());
                log.warn("NAV histórico incompleto en Valores_Fondo_Moder para inicio={}; se complementa con Consolidado de Rent_Vr_Uni_Moderado.", fechaInicio);
            }
        }
        IpcSeries ipcSeries = readIpcSeries(rentModeradoFile);
        return calcular(fechaInicio, fechaCorte, navSeries, ipcSeries);
    }

    private RentabilidadResultado calcular(
            LocalDate fechaInicio,
            LocalDate fechaFin,
            NavSeries navSeries,
            IpcSeries ipcSeries
    ) {
        NavigableMap<LocalDate, BigDecimal> nav = navSeries.values();
        NavigableMap<YearMonth, BigDecimal> ipc = ipcSeries.values();
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
        double ipcFactor = ipcFin.divide(ipcIni, 16, RoundingMode.HALF_UP).doubleValue();
        BigDecimal inflacionPeriodo = BigDecimal.valueOf(ipcFactor - 1d);
        BigDecimal inflacion = BigDecimal.valueOf(Math.pow(ipcFactor, 365d / (double) dias) - 1d);
        BigDecimal real = nominal.add(BigDecimal.ONE)
                .divide(inflacion.add(BigDecimal.ONE), 16, RoundingMode.HALF_UP)
                .subtract(BigDecimal.ONE);

        String navIniSrc = navSeries.sourceByDate().getOrDefault(navIniEntry.getKey(), "desconocido");
        String navFinSrc = navSeries.sourceByDate().getOrDefault(navFinEntry.getKey(), "desconocido");
        Integer navIniCount = navSeries.contributorsByDate().getOrDefault(navIniEntry.getKey(), 0);
        Integer navFinCount = navSeries.contributorsByDate().getOrDefault(navFinEntry.getKey(), 0);
        String ipcSource = "sheet=" + ipcSeries.sheetName() + (ipcSeries.convertedFromRates() ? " (tasas->índice)" : " (índice directo)");
        log.info("Rentabilidad detalle operandos: fechaInicio={} fechaFin={} dias={} | NAV_ini_fecha={} NAV_ini_valor={} NAV_ini_fuente={} NAV_ini_fondos={} | NAV_fin_fecha={} NAV_fin_valor={} NAV_fin_fuente={} NAV_fin_fondos={} | IPC_ini_mes={} IPC_ini_valor={} IPC_ini_fuente={} | IPC_fin_mes={} IPC_fin_valor={} IPC_fin_fuente={} | ipcFactor=(IPC_fin/IPC_ini)={} | inflacion_periodo=(IPC_fin/IPC_ini)-1={} | inflacion_anual=((IPC_fin/IPC_ini)^(365/dias))-1={} | nominal_anual=(NAV_fin/NAV_ini)^(365/dias)-1={} | real=((1+nominal_anual)/(1+inflacion_anual))-1={}",
                fechaInicio, fechaFin, dias,
                navIniEntry.getKey(), navIni, navIniSrc, navIniCount,
                navFinEntry.getKey(), navFin, navFinSrc, navFinCount,
                ipcIniEntry.getKey(), ipcIni, ipcSource,
                ipcFinEntry.getKey(), ipcFin, ipcSource,
                ipcFactor, inflacionPeriodo, inflacion, nominal, real);
        return new RentabilidadResultado(fechaInicio, fechaFin, nominal, real);
    }

    private NavSeries readNavPromedio(Path file, LocalDate fechaInicio, LocalDate fechaFin) {
            Map<LocalDate, List<BigDecimal>> porFecha = new TreeMap<>();
        Map<LocalDate, String> sourceByDate = new TreeMap<>();
        Map<LocalDate, Integer> contributorsByDate = new TreeMap<>();
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
                            sourceByDate.put(fecha, navFile.getFileName() + "/" + nombre + "!O");
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
                contributorsByDate.put(e.getKey(), e.getValue().size());
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
            return new NavSeries(serie, sourceByDate, contributorsByDate);
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

    private IpcSeries readIpcSeries(Path rentModeradoFile) {
        try (Workbook wb = WorkbookFactory.create(rentModeradoFile.toFile(), null, true)) {
            IpcSeries ipcBrSeries = null;
            Sheet ipcBr = getSheetIgnoreCase(wb, "IPC_BR");
            if (ipcBr != null) {
                NavigableMap<YearMonth, BigDecimal> serie = readDateValueSheetByMonth(ipcBr, 1, 2);
                if (!serie.isEmpty()) {
                    ipcBrSeries = new IpcSeries(serie, ipcBr.getSheetName(), false);
                    log.info("Serie IPC_BR cargada: file={} fechas={} desde={} hasta={}",
                            rentModeradoFile.toAbsolutePath(),
                            serie.size(),
                            serie.firstKey(),
                            serie.lastKey());
                }
            }

            IpcSeries ipcSeries = null;
            Sheet ipc = getSheetIgnoreCase(wb, "IPC");
            if (ipc != null) {
                NavigableMap<YearMonth, BigDecimal> tasas = readDateValueSheetByMonth(ipc, 1, 2);
                if (!tasas.isEmpty()) {
                    if (isIndexSeries(tasas)) {
                        log.info("Serie IPC cargada como índice directo: file={} fechas={}", rentModeradoFile.toAbsolutePath(), tasas.size());
                        ipcSeries = new IpcSeries(tasas, ipc.getSheetName(), false);
                    } else {
                        BigDecimal indice = BigDecimal.valueOf(100);
                        NavigableMap<YearMonth, BigDecimal> indices = new TreeMap<>();
                        for (var e : tasas.entrySet().stream().sorted(Map.Entry.comparingByKey()).collect(Collectors.toList())) {
                            indice = indice.multiply(BigDecimal.ONE.add(e.getValue())).setScale(16, RoundingMode.HALF_UP);
                            indices.put(e.getKey(), indice);
                        }
                        log.info("Serie IPC (tasas->índice) cargada: file={} fechas={}", rentModeradoFile.toAbsolutePath(), indices.size());
                        ipcSeries = new IpcSeries(indices, ipc.getSheetName(), true);
                    }
                }
            }

            // Selección defensiva: IPC_BR tiene prioridad solo si trae historia suficiente.
            if (ipcBrSeries != null && ipcBrSeries.values().size() >= 24) {
                return ipcBrSeries;
            }
            if (ipcBrSeries != null && ipcSeries != null) {
                log.warn("IPC_BR tiene cobertura corta ({} puntos); se usa IPC con mayor cobertura ({} puntos).",
                        ipcBrSeries.values().size(), ipcSeries.values().size());
                return ipcSeries;
            }
            if (ipcBrSeries != null) {
                return ipcBrSeries;
            }
            if (ipcSeries != null) {
                return ipcSeries;
            }
        } catch (Exception e) {
            log.warn("No fue posible leer IPC desde {}: {}", rentModeradoFile.toAbsolutePath(), e.getMessage());
        }
        return new IpcSeries(new TreeMap<>(), "N/A", false);
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

    private NavSeries readNavFromRentConsolidado(Path rentModeradoFile, LocalDate fechaFin) {
        NavigableMap<LocalDate, BigDecimal> data = new TreeMap<>();
        Map<LocalDate, String> sourceByDate = new TreeMap<>();
        Map<LocalDate, Integer> contributorsByDate = new TreeMap<>();
        try (Workbook wb = WorkbookFactory.create(rentModeradoFile.toFile(), null, true)) {
            Sheet s = getSheetIgnoreCase(wb, "Consolidado");
            if (s == null) return new NavSeries(data, sourceByDate, contributorsByDate);
            int last = s.getLastRowNum() + 1;
            for (int r = 14; r <= last; r++) { // estructura histórica conocida en consolidado
                Row row = s.getRow(r - 1);
                if (row == null) continue;
                LocalDate fecha = cellAsDate(row.getCell(0)); // col A
                BigDecimal nav = cellAsNumber(row.getCell(4)); // col E (NAV nominal usado por macro/tabla)
                if (fecha == null || nav.signum() <= 0) continue;
                if (fecha.isAfter(fechaFin)) continue;
                data.put(fecha, nav);
                sourceByDate.put(fecha, rentModeradoFile.getFileName() + "/Consolidado!E");
                contributorsByDate.put(fecha, 1);
            }
            log.info("Serie NAV desde Consolidado cargada: file={} fechas={} desde={} hasta={}",
                    rentModeradoFile.toAbsolutePath(),
                    data.size(),
                    data.isEmpty() ? null : data.firstKey(),
                    data.isEmpty() ? null : data.lastKey());
            return new NavSeries(data, sourceByDate, contributorsByDate);
        } catch (Exception e) {
            log.warn("No fue posible complementar NAV desde Consolidado en {}: {}", rentModeradoFile.toAbsolutePath(), e.getMessage());
            return new NavSeries(data, sourceByDate, contributorsByDate);
        }
    }

    private record NavSeries(
            NavigableMap<LocalDate, BigDecimal> values,
            Map<LocalDate, String> sourceByDate,
            Map<LocalDate, Integer> contributorsByDate
    ) {}

    private record IpcSeries(
            NavigableMap<YearMonth, BigDecimal> values,
            String sheetName,
            boolean convertedFromRates
    ) {}

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
