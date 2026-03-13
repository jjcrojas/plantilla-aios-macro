package co.gov.sfc.excel;

import co.gov.sfc.insumos.InsumosLocator;
import org.junit.jupiter.api.Test;

import java.math.BigDecimal;
import java.nio.file.Path;
import java.time.LocalDate;

import static org.junit.jupiter.api.Assertions.assertTrue;
import static org.mockito.Mockito.mock;
import static org.mockito.Mockito.when;

class TrimestralDataReaderTest {

    @Test
    void shouldReadTraspasosPerAfpFrom493UsingMacroCodes() {
        MensualDataReader mensualDataReader = mock(MensualDataReader.class);
        InsumosLocator locator = mock(InsumosLocator.class);

        LocalDate fecha = LocalDate.of(2025, 6, 30);
        when(locator.findRequired("493", fecha)).thenReturn(Path.of("insumos_ejemplo", "Serie_Formato_493 MOVIMIENTO AFILIADOS.xlsx"));
        when(mensualDataReader.read(fecha)).thenReturn(new MensualData(
                "jun-25", BigDecimal.ZERO, BigDecimal.ZERO, BigDecimal.ZERO,
                BigDecimal.ZERO, BigDecimal.ONE, BigDecimal.ZERO, BigDecimal.ZERO,
                BigDecimal.ZERO, BigDecimal.ZERO, BigDecimal.ZERO, BigDecimal.ZERO,
                BigDecimal.ZERO, BigDecimal.ZERO, BigDecimal.ZERO, BigDecimal.ZERO,
                BigDecimal.ZERO, BigDecimal.ZERO
        ));

        TrimestralDataReader reader = new TrimestralDataReader(mensualDataReader, locator);
        TrimestralData data = reader.read(fecha);

        assertTrue(data.traspasosColfondos().signum() > 0);
        assertTrue(data.traspasosPorvenir().signum() > 0);
        assertTrue(data.traspasosProteccion().signum() > 0);
        assertTrue(data.traspasosSkandia().signum() > 0);
    }
}
