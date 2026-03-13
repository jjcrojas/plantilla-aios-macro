package co.gov.sfc.services;

import co.gov.sfc.excel.MensualDataReader;
import co.gov.sfc.excel.MensualExcelGenerator;
import co.gov.sfc.excel.TrimestralData;
import co.gov.sfc.excel.TrimestralDataReader;
import co.gov.sfc.excel.TrimestralExcelGenerator;
import co.gov.sfc.model.ModoGeneracion;
import org.junit.jupiter.api.Test;

import java.math.BigDecimal;
import java.nio.file.Path;
import java.time.LocalDate;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.mockito.Mockito.mock;
import static org.mockito.Mockito.verify;
import static org.mockito.Mockito.when;

class AiosGeneracionServiceTest {

    @Test
    void shouldGenerateTrimestralWhenModeIsTrimestral() {
        MensualDataReader mensualDataReader = mock(MensualDataReader.class);
        MensualExcelGenerator mensualExcelGenerator = mock(MensualExcelGenerator.class);
        TrimestralDataReader trimestralDataReader = mock(TrimestralDataReader.class);
        TrimestralExcelGenerator trimestralExcelGenerator = mock(TrimestralExcelGenerator.class);

        AiosGeneracionService service = new AiosGeneracionService(
                mensualDataReader,
                mensualExcelGenerator,
                trimestralDataReader,
                trimestralExcelGenerator
        );

        LocalDate fecha = LocalDate.of(2025, 6, 30);
        TrimestralData data = new TrimestralData("jun-25", BigDecimal.ONE, BigDecimal.ONE, BigDecimal.ONE, BigDecimal.ONE,
                BigDecimal.ONE, BigDecimal.ONE, BigDecimal.ONE, BigDecimal.ONE, BigDecimal.ONE, BigDecimal.ONE, BigDecimal.ONE);
        when(trimestralDataReader.read(fecha)).thenReturn(data);
        when(trimestralExcelGenerator.generar(fecha, data)).thenReturn(Path.of("target/aios-output/Boletin_AIOS TRIMESTRAL.xlsx"));

        var resultado = service.generar(fecha, ModoGeneracion.TRIMESTRAL);

        assertEquals(1, resultado.archivosGenerados().size());
        assertEquals("Boletin_AIOS TRIMESTRAL.xlsx", resultado.archivosGenerados().getFirst().getFileName().toString());
        verify(trimestralDataReader).read(fecha);
        verify(trimestralExcelGenerator).generar(fecha, data);
    }

    @Test
    void shouldRejectTrimestralForNonQuarterMonth() {
        AiosGeneracionService service = new AiosGeneracionService(
                mock(MensualDataReader.class),
                mock(MensualExcelGenerator.class),
                mock(TrimestralDataReader.class),
                mock(TrimestralExcelGenerator.class)
        );

        IllegalArgumentException ex = assertThrows(IllegalArgumentException.class,
                () -> service.generar(LocalDate.of(2025, 5, 31), ModoGeneracion.TRIMESTRAL));

        assertEquals("La generación trimestral solo aplica para cortes de marzo, junio, septiembre o diciembre", ex.getMessage());
    }
}
