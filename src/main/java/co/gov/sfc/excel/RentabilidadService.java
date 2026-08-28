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
import java.util.Map;
import java.util.NavigableMap;
import java.util.TreeMap;

@Service
public class RentabilidadService {

    private static final Logger log = LoggerFactory.getLogger(RentabilidadService.class);

    /**
     * Calcula las rentabilidades con las mismas fuentes del libro original:
     * Consolidado!E para el NAV sintético e IPC_D!B para el índice diario.
     */
    public RentabilidadResultado calcularRentabilidad(
            Path rentModeradoFile,
            LocalDate fechaCorte,
            int horizonteAnios
    ) {
        if (horizonteAnios <= 0) {
            throw new IllegalArgumentException("El horizonte debe ser mayor que cero: " + horizonteAnios);
        }
        LocalDate fechaInicio = fechaCorte.minusYears(horizonteAnios);
        RentabilidadSeries series = readRentabilidadSeries(rentModeradoFile, fechaCorte);
        return calcular(fechaInicio, fechaCorte, horizonteAnios, series.nav(), series.ipcDiario());
    }

    /**
     * Conserva compatibilidad con llamadas anteriores. Valores_Fondo_Moder ya no participa
     * en el cálculo porque su promedio por AFP no equivale al NAV sintético de Consolidado!E.
     */
    public RentabilidadResultado calcularRentabilidad(
            Path valoresFondoModerFile,
            Path rentModeradoFile,
            LocalDate fechaCorte,
            int horizonteAnios
    ) {
        log.debug("Se omite Valores_Fondo_Moder={} para rentabilidades; fuente NAV obligatoria={}/Consolidado!E.",
                valoresFondoModerFile, rentModeradoFile.toAbsolutePath());
        return calcularRentabilidad(rentModeradoFile, fechaCorte, horizonteAnios);
    }

    private RentabilidadResultado calcular(
            LocalDate fechaInicio,
            LocalDate fechaFin,
            int horizonteAnios,
            NavSeries navSeries,
            IpcDailySeries ipcSeries
    ) {
        NavObservation navIni = requireNav(navSeries, fechaInicio, "inicial");
        NavObservation navFin = requireNav(navSeries, fechaFin, "final");
        IpcObservation ipcIni = requireIpc(ipcSeries, fechaInicio, "inicial");
        IpcObservation ipcFin = requireIpc(ipcSeries, fechaFin, "final");

        validarPositivo(navIni.value(), "NAV inicial", fechaInicio);
        validarPositivo(navFin.value(), "NAV final", fechaFin);
        validarPositivo(ipcIni.value(), "IPC inicial", fechaInicio);
        validarPositivo(ipcFin.value(), "IPC final", fechaFin);

        long dias = ChronoUnit.DAYS.between(fechaInicio, fechaFin);
        if (dias <= 0) {
            throw new IllegalStateException("Rango inválido para rentabilidad: " + fechaInicio + " a " + fechaFin);
        }

        double navFactor = navFin.value().divide(navIni.value(), 16, RoundingMode.HALF_UP).doubleValue();
        double ipcFactor = ipcFin.value().divide(ipcIni.value(), 16, RoundingMode.HALF_UP).doubleValue();
        double exponenteAnualizacion = 365d / (double) dias;
        BigDecimal nominal = BigDecimal.valueOf(Math.pow(navFactor, exponenteAnualizacion) - 1d);
        BigDecimal real = BigDecimal.valueOf(Math.pow(navFactor / ipcFactor, exponenteAnualizacion) - 1d);

        log.info("Rentabilidad auditoría horizonte={} años: fechaInicioSolicitada={} fechaInicioUsada={} "
                        + "NAV_inicial={} ubicaciónNAVInicial={} IPC_inicial={} ubicaciónIPCInicial={} | "
                        + "fechaFinSolicitada={} fechaFinUsada={} NAV_final={} ubicaciónNAVFinal={} "
                        + "IPC_final={} ubicaciónIPCFinal={}",
                horizonteAnios,
                fechaInicio, navIni.date(), navIni.value(), formatNavObservation(navIni),
                ipcIni.value(), formatIpcObservation(ipcIni),
                fechaFin, navFin.date(), navFin.value(), formatNavObservation(navFin),
                ipcFin.value(), formatIpcObservation(ipcFin));
        log.info("Rentabilidad resultado horizonte={} años: dias={} exponenteAnualizacion={} "
                        + "factorNAV={} factorIPC={} nominal={} real={} fuentes=Consolidado!E+IPC_D!B",
                horizonteAnios, dias, exponenteAnualizacion, navFactor, ipcFactor, nominal, real);

        return new RentabilidadResultado(fechaInicio, fechaFin, nominal, real);
    }

    private RentabilidadSeries readRentabilidadSeries(Path rentModeradoFile, LocalDate fechaFin) {
        try (Workbook wb = WorkbookFactory.create(rentModeradoFile.toFile(), null, true)) {
            NavSeries nav = readNavFromRentConsolidado(wb, rentModeradoFile, fechaFin);
            IpcDailySeries ipc = readIpcDiario(wb, rentModeradoFile, fechaFin);
            if (nav.values().isEmpty()) {
                throw new IllegalStateException("No hay valores válidos en Consolidado!E");
            }
            if (ipc.values().isEmpty()) {
                throw new IllegalStateException("No hay valores válidos en IPC_D!B");
            }
            log.info("Fuentes de rentabilidad cargadas: file={} NAV=Consolidado!E puntos={} desde={} hasta={} "
                            + "IPC=IPC_D!B puntos={} desde={} hasta={}",
                    rentModeradoFile.toAbsolutePath(),
                    nav.values().size(), nav.values().firstKey(), nav.values().lastKey(),
                    ipc.values().size(), ipc.values().firstKey(), ipc.values().lastKey());
            return new RentabilidadSeries(nav, ipc);
        } catch (IllegalStateException e) {
            throw e;
        } catch (Exception e) {
            throw new IllegalStateException("No fue posible leer rentabilidades desde "
                    + rentModeradoFile.toAbsolutePath() + ": " + e.getMessage(), e);
        }
    }

    private NavSeries readNavFromRentConsolidado(Workbook wb, Path file, LocalDate fechaFin) {
        Sheet sheet = getSheetIgnoreCase(wb, "Consolidado");
        if (sheet == null) {
            throw new IllegalStateException("No existe la hoja Consolidado en " + file.toAbsolutePath());
        }
        NavigableMap<LocalDate, BigDecimal> values = new TreeMap<>();
        Map<LocalDate, NavObservation> observations = new TreeMap<>();
        int last = sheet.getLastRowNum() + 1;
        for (int r = 14; r <= last; r++) {
            Row row = sheet.getRow(r - 1);
            if (row == null) continue;
            LocalDate fecha = cellAsDate(row.getCell(0));
            BigDecimal nav = cellAsNumber(row.getCell(4));
            if (fecha == null || nav.signum() <= 0 || fecha.isAfter(fechaFin)) continue;
            values.put(fecha, nav);
            observations.put(fecha, new NavObservation(
                    fecha, file.toAbsolutePath().toString(), sheet.getSheetName(), "A" + r, "E" + r, nav));
        }
        return new NavSeries(values, observations);
    }

    private IpcDailySeries readIpcDiario(Workbook wb, Path file, LocalDate fechaFin) {
        Sheet sheet = getSheetIgnoreCase(wb, "IPC_D");
        if (sheet == null) {
            throw new IllegalStateException("No existe la hoja IPC_D en " + file.toAbsolutePath());
        }
        NavigableMap<LocalDate, BigDecimal> values = new TreeMap<>();
        Map<LocalDate, IpcObservation> observations = new TreeMap<>();
        int last = sheet.getLastRowNum() + 1;
        for (int r = 2; r <= last; r++) {
            Row row = sheet.getRow(r - 1);
            if (row == null) continue;
            LocalDate fecha = cellAsDate(row.getCell(0));
            BigDecimal ipc = cellAsNumber(row.getCell(1));
            if (fecha == null || ipc.signum() <= 0 || fecha.isAfter(fechaFin)) continue;
            values.put(fecha, ipc);
            observations.put(fecha, new IpcObservation(
                    fecha, file.toAbsolutePath().toString(), sheet.getSheetName(), "A" + r, "B" + r, ipc));
        }
        return new IpcDailySeries(values, observations);
    }

    private NavObservation requireNav(NavSeries series, LocalDate date, String extremo) {
        NavObservation observation = series.observations().get(date);
        if (observation == null) {
            throw new IllegalStateException("No existe NAV " + extremo + " para la fecha exacta " + date
                    + " en Consolidado!E; cobertura=" + coverage(series.values()));
        }
        return observation;
    }

    private IpcObservation requireIpc(IpcDailySeries series, LocalDate date, String extremo) {
        IpcObservation observation = series.observations().get(date);
        if (observation == null) {
            throw new IllegalStateException("No existe IPC diario " + extremo + " para la fecha exacta " + date
                    + " en IPC_D!B; cobertura=" + coverage(series.values()));
        }
        return observation;
    }

    private String coverage(NavigableMap<LocalDate, BigDecimal> values) {
        return values.isEmpty() ? "vacía" : values.firstKey() + " a " + values.lastKey();
    }

    private void validarPositivo(BigDecimal value, String nombre, LocalDate fecha) {
        if (value == null || value.signum() <= 0) {
            throw new IllegalStateException(nombre + " inválido para " + fecha + ": " + value);
        }
    }

    private String formatNavObservation(NavObservation observation) {
        return "{archivo=" + observation.file() + ", hoja=" + observation.sheet()
                + ", fecha=" + observation.dateCell() + ", NAV=" + observation.valueCell()
                + ", valor=" + observation.value() + "}";
    }

    private String formatIpcObservation(IpcObservation observation) {
        return "{archivo=" + observation.file() + ", hoja=" + observation.sheet()
                + ", fecha=" + observation.dateCell() + ", IPC=" + observation.valueCell()
                + ", valor=" + observation.value() + "}";
    }

    private LocalDate cellAsDate(Cell cell) {
        if (cell == null) return null;
        try {
            return switch (cell.getCellType()) {
                case NUMERIC -> cell.getDateCellValue().toInstant().atZone(ZoneId.systemDefault()).toLocalDate();
                case STRING -> {
                    String value = cell.getStringCellValue();
                    yield value == null || value.isBlank() ? null : LocalDate.parse(value.trim());
                }
                case FORMULA -> cell.getCachedFormulaResultType() == org.apache.poi.ss.usermodel.CellType.NUMERIC
                        ? cell.getDateCellValue().toInstant().atZone(ZoneId.systemDefault()).toLocalDate()
                        : null;
                default -> null;
            };
        } catch (Exception ignore) {
            return null;
        }
    }

    private BigDecimal cellAsNumber(Cell cell) {
        if (cell == null) return BigDecimal.ZERO;
        try {
            return switch (cell.getCellType()) {
                case NUMERIC -> BigDecimal.valueOf(cell.getNumericCellValue());
                case FORMULA -> cell.getCachedFormulaResultType() == org.apache.poi.ss.usermodel.CellType.NUMERIC
                        ? BigDecimal.valueOf(cell.getNumericCellValue())
                        : BigDecimal.ZERO;
                case STRING -> parseDecimal(cell.getStringCellValue());
                default -> BigDecimal.ZERO;
            };
        } catch (Exception ignore) {
            return BigDecimal.ZERO;
        }
    }

    private BigDecimal parseDecimal(String value) {
        if (value == null || value.isBlank()) return BigDecimal.ZERO;
        try {
            return new BigDecimal(value.trim().replace(",", ""));
        } catch (Exception ignore) {
            return BigDecimal.ZERO;
        }
    }

    private Sheet getSheetIgnoreCase(Workbook wb, String name) {
        for (int i = 0; i < wb.getNumberOfSheets(); i++) {
            Sheet sheet = wb.getSheetAt(i);
            if (sheet.getSheetName().equalsIgnoreCase(name)) return sheet;
        }
        return null;
    }

    private record RentabilidadSeries(NavSeries nav, IpcDailySeries ipcDiario) {}

    private record NavSeries(
            NavigableMap<LocalDate, BigDecimal> values,
            Map<LocalDate, NavObservation> observations
    ) {}

    private record IpcDailySeries(
            NavigableMap<LocalDate, BigDecimal> values,
            Map<LocalDate, IpcObservation> observations
    ) {}

    private record NavObservation(
            LocalDate date,
            String file,
            String sheet,
            String dateCell,
            String valueCell,
            BigDecimal value
    ) {}

    private record IpcObservation(
            LocalDate date,
            String file,
            String sheet,
            String dateCell,
            String valueCell,
            BigDecimal value
    ) {}

    public record RentabilidadResultado(
            LocalDate fechaInicio,
            LocalDate fechaFin,
            BigDecimal rentabilidadNominal,
            BigDecimal rentabilidadReal
    ) {}
}
