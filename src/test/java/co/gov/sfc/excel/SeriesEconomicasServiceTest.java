package co.gov.sfc.excel;

import co.gov.sfc.insumos.InsumosLocator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.math.BigDecimal;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.LocalDate;
import java.time.ZoneId;
import java.util.Date;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.mockito.Mockito.mock;
import static org.mockito.Mockito.when;

class SeriesEconomicasServiceTest {

    @TempDir
    Path tempDir;

    @Test
    void shouldReadPeaPibAndLatestPositiveGovernmentDebtFromHoja1() throws Exception {
        Path archivo = tempDir.resolve("series PIB_PEA_TRM_DG.xlsm");
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            Sheet hoja = workbook.createSheet("Hoja1");
            Row anterior = hoja.createRow(1);
            ponerSerie(anterior, 5, 6, LocalDate.of(2024, 12, 31), 1000);
            ponerSerie(anterior, 8, 9, LocalDate.of(2024, 12, 31), 310);
            ponerSerie(anterior, 11, 12, LocalDate.of(2024, 12, 31), 500);

            Row corte = hoja.createRow(2);
            ponerSerie(corte, 5, 6, LocalDate.of(2025, 6, 30), 1100);
            ponerSerie(corte, 8, 9, LocalDate.of(2025, 6, 30), 320);
            ponerSerie(corte, 11, 12, LocalDate.of(2025, 6, 30), 0);
            try (var out = Files.newOutputStream(archivo)) {
                workbook.write(out);
            }
        }
        InsumosLocator locator = mock(InsumosLocator.class);
        LocalDate fechaCorte = LocalDate.of(2025, 6, 30);
        when(locator.findRequired("PIB_PEA_TRM_DG", fechaCorte)).thenReturn(archivo);

        SeriesEconomicasService.SeriesEconomicas series =
                new SeriesEconomicasService(locator).leer(fechaCorte);

        assertEquals(new BigDecimal("320.0"), series.pea());
        assertEquals(new BigDecimal("1100.0"), series.pibSemestral());
        assertEquals(new BigDecimal("500.0"), series.deudaGubernamental());
        assertEquals(archivo.toAbsolutePath(), series.archivo());
    }

    private void ponerSerie(Row row, int columnaFecha, int columnaValor, LocalDate fecha, double valor) {
        Date fechaExcel = Date.from(fecha.atStartOfDay(ZoneId.systemDefault()).toInstant());
        row.createCell(columnaFecha).setCellValue(fechaExcel);
        row.createCell(columnaValor).setCellValue(valor);
    }
}
