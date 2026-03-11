package co.gov.sfc.insumos;

import co.gov.sfc.config.AiosProperties;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.nio.file.Files;
import java.nio.file.Path;
import java.time.LocalDate;

import static org.junit.jupiter.api.Assertions.assertEquals;

class InsumosLocatorTest {

    @TempDir
    Path tempDir;

    @Test
    void shouldResolveSistemaTotalInMacroBalancesStructure() throws Exception {
        Path monthDir = tempDir.resolve("Balances/2025/6 junio");
        Files.createDirectories(monthDir);
        Path file = monthDir.resolve("SISTEMA TOTAL junio 2025.xls");
        Files.createFile(file);

        InsumosLocator locator = new InsumosLocator(new AiosProperties(tempDir, tempDir, tempDir));

        Path resolved = locator.findRequired("SISTEMA TOTAL", LocalDate.of(2025, 6, 30));

        assertEquals(file, resolved);
    }

    @Test
    void shouldResolveFormato491InMacroSubfolder() throws Exception {
        Path dir491 = tempDir.resolve("Formato 491/491 FORMATO TRANSMITIDO");
        Files.createDirectories(dir491);
        Path file = dir491.resolve("Serie_Formato_ 491 AFILIADOS AFP.xlsm");
        Files.createFile(file);

        InsumosLocator locator = new InsumosLocator(new AiosProperties(tempDir, tempDir, tempDir));

        Path resolved = locator.findRequired("491", LocalDate.of(2025, 6, 30));

        assertEquals(file, resolved);
    }


    @Test
    void shouldResolveRentabilidadModeradoInMacroYearMonthFolder() throws Exception {
        Path rentDir = tempDir.resolve("PROCESOS MENSUALES/Rentabilidad Minima/Historico_Rent_minima/2025/6 Junio");
        Files.createDirectories(rentDir);
        Path file = rentDir.resolve("Rent_Vr_Uni_Moderado.xlsm");
        Files.createFile(file);

        InsumosLocator locator = new InsumosLocator(new AiosProperties(tempDir, tempDir, tempDir));

        Path resolved = locator.findRequired("Rent_Vr_Uni_Moderado", LocalDate.of(2025, 6, 30));

        assertEquals(file, resolved);
    }

    @Test
    void shouldResolveLimitesInMacroYearMonthFolder() throws Exception {
        Path limitesDir = tempDir.resolve("LIMITES/2025/6 Junio");
        Files.createDirectories(limitesDir);
        Path file = limitesDir.resolve("LIMITES del nuevo.xlsm");
        Files.createFile(file);

        InsumosLocator locator = new InsumosLocator(new AiosProperties(tempDir, tempDir, tempDir));

        Path resolved = locator.findRequired("LIMITES", LocalDate.of(2025, 6, 30));

        assertEquals(file, resolved);
    }

}
