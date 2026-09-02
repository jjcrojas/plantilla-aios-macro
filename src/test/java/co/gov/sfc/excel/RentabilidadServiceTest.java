package co.gov.sfc.excel;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.LocalDate;
import java.util.LinkedHashMap;
import java.util.Map;

import static org.junit.jupiter.api.Assertions.assertEquals;

class RentabilidadServiceTest {

    @Test
    void shouldMatchExcelForAllHorizonsUsingConsolidadoAndDailyIpc() throws Exception {
        Path rentFile = createRentabilidadWorkbook("todos-horizontes");
        RentabilidadService service = new RentabilidadService();
        LocalDate fechaCorte = LocalDate.of(2025, 6, 30);
        Map<Integer, Expected> expectedByYears = new LinkedHashMap<>();
        expectedByYears.put(10, new Expected(0.08803839124002377, 0.028058005172205247));
        expectedByYears.put(5, new Expected(0.10078464820855304, 0.02456775183694515));
        expectedByYears.put(3, new Expected(0.11165268194584987, 0.029371740277118086));
        expectedByYears.put(1, new Expected(0.1022608588509073, 0.051511390166620874));

        for (var entry : expectedByYears.entrySet()) {
            RentabilidadService.RentabilidadResultado result = service.calcularRentabilidad(
                    rentFile, fechaCorte, entry.getKey());

            assertEquals(entry.getValue().nominal(), result.rentabilidadNominal().doubleValue(), 1e-12,
                    "nominal " + entry.getKey() + " años");
            assertEquals(entry.getValue().real(), result.rentabilidadReal().doubleValue(), 1e-12,
                    "real " + entry.getKey() + " años");
        }
    }

    @Test
    void shouldNotRequireLegacyValoresFondoModerFile() throws Exception {
        Path rentFile = createRentabilidadWorkbook("sin-valores-fondo");
        Path legacyFileThatDoesNotExist = rentFile.getParent().resolve("Valores_Fondo_Moder-inexistente.xlsx");

        RentabilidadService.RentabilidadResultado result = new RentabilidadService().calcularRentabilidad(
                legacyFileThatDoesNotExist, rentFile, LocalDate.of(2025, 6, 30), 1);

        assertEquals(0.1022608588509073, result.rentabilidadNominal().doubleValue(), 1e-12);
        assertEquals(0.051511390166620874, result.rentabilidadReal().doubleValue(), 1e-12);
    }

    private Path createRentabilidadWorkbook(String name) throws Exception {
        Path directory = Path.of("target", "test-rentabilidad");
        Files.createDirectories(directory);
        Path file = directory.resolve("Rent_Vr_Uni_Moderado-" + name + ".xlsx");
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            Sheet consolidado = workbook.createSheet("Consolidado");
            Sheet ipcDiario = workbook.createSheet("IPC_D");
            ipcDiario.createRow(0).createCell(1).setCellValue("IPC diario");

            LocalDate[] dates = {
                    LocalDate.of(2015, 6, 30),
                    LocalDate.of(2020, 6, 30),
                    LocalDate.of(2022, 6, 30),
                    LocalDate.of(2024, 6, 30),
                    LocalDate.of(2025, 6, 30)
            };
            double[] navValues = {
                    429790.30857641913,
                    618548.7393695777,
                    727723.9982261503,
                    907226.2631574227,
                    1000000
            };
            double[] ipcValues = {
                    85.20999999999991,
                    104.9700000000001,
                    119.31000000000019,
                    143.38000000000008,
                    150.30000000000007
            };

            for (int i = 0; i < dates.length; i++) {
                Row navRow = consolidado.createRow(13 + i);
                navRow.createCell(0).setCellValue(java.sql.Date.valueOf(dates[i]));
                navRow.createCell(4).setCellValue(navValues[i]);
                if (dates[i].equals(LocalDate.of(2025, 6, 30))) {
                    // Valor deliberadamente incorrecto para demostrar que no se lee Consolidado!I.
                    navRow.createCell(8).setCellValue(0.0011923327597287425);
                }

                Row ipcRow = ipcDiario.createRow(1 + i);
                ipcRow.createCell(0).setCellValue(java.sql.Date.valueOf(dates[i]));
                ipcRow.createCell(1).setCellValue(ipcValues[i]);
            }

            try (OutputStream output = Files.newOutputStream(file)) {
                workbook.write(output);
            }
        }
        return file;
    }

    private record Expected(double nominal, double real) {}
}
