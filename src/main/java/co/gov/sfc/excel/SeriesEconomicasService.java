package co.gov.sfc.excel;

import co.gov.sfc.insumos.InsumosLocator;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Component;

import java.math.BigDecimal;
import java.nio.file.Path;
import java.time.LocalDate;
import java.time.ZoneId;

@Component
public class SeriesEconomicasService {

    private static final Logger log = LoggerFactory.getLogger(SeriesEconomicasService.class);
    private final InsumosLocator locator;

    public SeriesEconomicasService(InsumosLocator locator) {
        this.locator = locator;
    }

    public SeriesEconomicas leer(LocalDate fechaCorte) {
        Path archivo = locator.findRequired("PIB_PEA_TRM_DG", fechaCorte);
        try (Workbook workbook = WorkbookFactory.create(archivo.toFile(), null, true)) {
            Sheet hoja = workbook.getSheet("Hoja1");
            if (hoja == null) {
                throw new IllegalStateException("No existe la hoja Hoja1 en " + archivo.toAbsolutePath());
            }
            DatoSerie pea = buscarUltimoValor(hoja, fechaCorte, 8, 9, "PEA");
            DatoSerie pib = buscarUltimoValor(hoja, fechaCorte, 5, 6, "PIB semestral");
            DatoSerie deuda = buscarUltimoValor(hoja, fechaCorte, 11, 12, "deuda gubernamental");
            log.info("Series económicas cargadas archivo={} fechaCorte={} PEA={} fechaPEA={} PIB={} fechaPIB={} deudaGubernamental={} fechaDeuda={}",
                    archivo.toAbsolutePath(), fechaCorte,
                    pea.valor(), pea.fecha(), pib.valor(), pib.fecha(), deuda.valor(), deuda.fecha());
            return new SeriesEconomicas(pea.valor(), pib.valor(), deuda.valor(), archivo.toAbsolutePath());
        } catch (Exception e) {
            throw new IllegalStateException("No fue posible leer PEA, PIB y deuda gubernamental desde "
                    + archivo.toAbsolutePath(), e);
        }
    }

    private DatoSerie buscarUltimoValor(Sheet hoja,
                                        LocalDate fechaCorte,
                                        int columnaFecha,
                                        int columnaValor,
                                        String nombreSerie) {
        DatoSerie mejor = null;
        for (Row row : hoja) {
            LocalDate fecha = fecha(row.getCell(columnaFecha));
            BigDecimal valor = numero(row.getCell(columnaValor));
            if (fecha == null || fecha.isAfter(fechaCorte) || valor.signum() <= 0) continue;
            DatoSerie actual = new DatoSerie(fecha, valor, row.getRowNum() + 1);
            if (fecha.equals(fechaCorte)) return actual;
            if (mejor == null || fecha.isAfter(mejor.fecha())) mejor = actual;
        }
        if (mejor == null) {
            throw new IllegalStateException("No se encontró " + nombreSerie + " para " + fechaCorte
                    + " ni para una fecha anterior en Hoja1");
        }
        log.info("Serie {} usa fecha anterior fechaCorte={} fechaSerie={} fila={} valor={}",
                nombreSerie, fechaCorte, mejor.fecha(), mejor.fila(), mejor.valor());
        return mejor;
    }

    private LocalDate fecha(Cell cell) {
        if (cell == null) return null;
        try {
            if (cell.getCellType() == CellType.NUMERIC ||
                    (cell.getCellType() == CellType.FORMULA
                            && cell.getCachedFormulaResultType() == CellType.NUMERIC)) {
                return cell.getDateCellValue().toInstant().atZone(ZoneId.systemDefault()).toLocalDate();
            }
        } catch (Exception ignored) {
            return null;
        }
        return null;
    }

    private BigDecimal numero(Cell cell) {
        if (cell == null) return BigDecimal.ZERO;
        try {
            if (cell.getCellType() == CellType.NUMERIC ||
                    (cell.getCellType() == CellType.FORMULA
                            && cell.getCachedFormulaResultType() == CellType.NUMERIC)) {
                return BigDecimal.valueOf(cell.getNumericCellValue());
            }
        } catch (Exception ignored) {
            return BigDecimal.ZERO;
        }
        return BigDecimal.ZERO;
    }

    private record DatoSerie(LocalDate fecha, BigDecimal valor, int fila) {}

    public record SeriesEconomicas(
            BigDecimal pea,
            BigDecimal pibSemestral,
            BigDecimal deudaGubernamental,
            Path archivo
    ) {}
}
