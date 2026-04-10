package co.gov.sfc.services;

import co.gov.sfc.excel.MensualData;
import co.gov.sfc.excel.MensualDataReader;
import co.gov.sfc.excel.MensualExcelGenerator;
import co.gov.sfc.excel.SemestralExcelGenerator;
import co.gov.sfc.excel.TrimestralData;
import co.gov.sfc.excel.TrimestralDataReader;
import co.gov.sfc.excel.TrimestralExcelGenerator;
import co.gov.sfc.model.ModoGeneracion;
import org.junit.jupiter.api.Test;

import java.nio.file.Path;
import java.time.LocalDate;
import java.util.Map;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.mockito.Mockito.*;

class AiosGeneracionServiceTest {

    @Test
    void shouldGenerateTrimestralWhenModeIsTrimestral() {
        MensualDataReader mensualDataReader = mock(MensualDataReader.class);
        MensualExcelGenerator mensualExcelGenerator = mock(MensualExcelGenerator.class);
        TrimestralDataReader trimestralDataReader = mock(TrimestralDataReader.class);
        TrimestralExcelGenerator trimestralExcelGenerator = mock(TrimestralExcelGenerator.class);
        SemestralExcelGenerator semestralExcelGenerator = mock(SemestralExcelGenerator.class);

        AiosGeneracionService service = new AiosGeneracionService(mensualDataReader, mensualExcelGenerator, semestralExcelGenerator, trimestralDataReader, trimestralExcelGenerator);

        LocalDate fecha = LocalDate.of(2025, 6, 30);
        TrimestralData data = new TrimestralData("jun-25", Map.of(), Map.of(), Map.of(), Map.of(), Map.of(), Map.of(), Map.of(), Map.of());
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
                mock(SemestralExcelGenerator.class),
                mock(TrimestralDataReader.class),
                mock(TrimestralExcelGenerator.class)
        );

        IllegalArgumentException ex = assertThrows(IllegalArgumentException.class,
                () -> service.generar(LocalDate.of(2025, 5, 31), ModoGeneracion.TRIMESTRAL));

        assertEquals("La generación trimestral solo aplica para cortes de marzo, junio, septiembre o diciembre", ex.getMessage());
    }

    @Test
    void shouldGenerateSemestralWhenModeIsSemestral() {
        MensualDataReader mensualDataReader = mock(MensualDataReader.class);
        MensualExcelGenerator mensualExcelGenerator = mock(MensualExcelGenerator.class);
        TrimestralDataReader trimestralDataReader = mock(TrimestralDataReader.class);
        TrimestralExcelGenerator trimestralExcelGenerator = mock(TrimestralExcelGenerator.class);
        SemestralExcelGenerator semestralExcelGenerator = mock(SemestralExcelGenerator.class);

        AiosGeneracionService service = new AiosGeneracionService(mensualDataReader, mensualExcelGenerator, semestralExcelGenerator, trimestralDataReader, trimestralExcelGenerator);

        LocalDate fecha = LocalDate.of(2025, 6, 30);
        java.math.BigDecimal one = java.math.BigDecimal.ONE;
        MensualData mensual = new MensualData("jun-25",
                one, one, one, one, one, one,
                one, one, one, one, one, one,
                one, one, one, one, one, one,
                one, one, one, one, one, one,
                one, one, one, one, one, one,
                one, one, one, one, one, one);
        TrimestralData data = new TrimestralData("jun-25", Map.of(), Map.of(), Map.of(), Map.of(), Map.of(), Map.of(), Map.of(), Map.of());
        when(mensualDataReader.read(fecha)).thenReturn(mensual);
        when(trimestralDataReader.read(fecha, mensual)).thenReturn(data);
        when(semestralExcelGenerator.generar(fecha, mensual, data)).thenReturn(Path.of("target/aios-output/semestral.xlsx"));

        var resultado = service.generar(fecha, ModoGeneracion.SEMESTRAL);

        assertEquals(1, resultado.archivosGenerados().size());
        assertEquals("semestral.xlsx", resultado.archivosGenerados().getFirst().getFileName().toString());
        verify(mensualDataReader).read(fecha);
        verify(trimestralDataReader).read(fecha, mensual);
        verify(semestralExcelGenerator).generar(fecha, mensual, data);
    }

    @Test
    void shouldRejectSemestralForNonSemesterMonth() {
        AiosGeneracionService service = new AiosGeneracionService(
                mock(MensualDataReader.class),
                mock(MensualExcelGenerator.class),
                mock(SemestralExcelGenerator.class),
                mock(TrimestralDataReader.class),
                mock(TrimestralExcelGenerator.class)
        );

        IllegalArgumentException ex = assertThrows(IllegalArgumentException.class,
                () -> service.generar(LocalDate.of(2025, 9, 30), ModoGeneracion.SEMESTRAL));

        assertEquals("La generación semestral solo aplica para cortes de junio o diciembre", ex.getMessage());
    }
}
