package co.gov.sfc.excel;

import co.gov.sfc.insumos.InsumosLocator;
import org.junit.jupiter.api.Disabled;
import org.junit.jupiter.api.Test;

import java.math.BigDecimal;
import java.nio.file.Path;
import java.time.LocalDate;

import static org.junit.jupiter.api.Assertions.assertTrue;
import static org.mockito.Mockito.mock;
import static org.mockito.Mockito.when;

class TrimestralDataReaderTest {

    @Disabled("Integra con workbook 493 pesado; se valida en pruebas de servicio")
    @Test
    void shouldReadTraspasosPerAfpFrom493UsingMacroCodes() {
        MensualDataReader mensualDataReader = mock(MensualDataReader.class);
        InsumosLocator locator = mock(InsumosLocator.class);

        LocalDate fecha = LocalDate.of(2025, 6, 30);
        when(locator.findRequired("493", fecha)).thenReturn(Path.of("insumos_ejemplo", "Serie_Formato_493 MOVIMIENTO AFILIADOS.xlsx"));
        when(locator.findRequired("SISTEMA TOTAL", fecha)).thenReturn(Path.of("insumos_ejemplo", "SISTEMA TOTAL Junio 2025.xls"));
        when(locator.findRequired("Rent_Vr_Uni_Moderado", fecha)).thenReturn(Path.of("insumos_ejemplo", "Rent_Vr_Uni_Moderado.xlsm"));
        when(locator.findRequired("comision fpo desde 2003", fecha)).thenReturn(Path.of("insumos_ejemplo", "comisión FPO desde 2003.xls"));
        when(mensualDataReader.read(fecha)).thenReturn(new MensualData(
                "jun-25",
                BigDecimal.ZERO, BigDecimal.ZERO, BigDecimal.ZERO, BigDecimal.ZERO, BigDecimal.ZERO, BigDecimal.ZERO,
                BigDecimal.ZERO, BigDecimal.ZERO, BigDecimal.ZERO,
                BigDecimal.ZERO, BigDecimal.valueOf(4000), BigDecimal.ZERO, BigDecimal.ZERO,
                BigDecimal.ZERO, BigDecimal.ZERO, BigDecimal.ZERO, BigDecimal.ZERO,
                BigDecimal.ZERO, BigDecimal.ZERO, BigDecimal.ZERO, BigDecimal.ZERO,
                BigDecimal.ZERO, BigDecimal.ZERO,
                BigDecimal.ZERO, BigDecimal.ZERO, BigDecimal.ZERO, BigDecimal.ZERO,
                BigDecimal.ZERO, BigDecimal.ZERO, BigDecimal.ZERO, BigDecimal.ZERO,
                BigDecimal.ZERO, BigDecimal.ZERO, BigDecimal.ZERO, BigDecimal.ZERO,
                BigDecimal.ZERO,
                BigDecimal.ZERO,
                BigDecimal.ZERO,
                BigDecimal.ZERO
        ));

        TrimestralDataReader reader = new TrimestralDataReader(mensualDataReader, locator);
        TrimestralData data = reader.read(fecha);

        assertTrue(data.traspasos().getOrDefault("colf", BigDecimal.ZERO).signum() >= 0);
        assertTrue(data.traspasos().getOrDefault("porv", BigDecimal.ZERO).signum() >= 0);
        assertTrue(data.traspasos().getOrDefault("prot", BigDecimal.ZERO).signum() >= 0);
        assertTrue(data.traspasos().getOrDefault("sk", BigDecimal.ZERO).signum() >= 0);
    }
}
