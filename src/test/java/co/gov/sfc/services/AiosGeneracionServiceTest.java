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
import org.junit.jupiter.api.io.TempDir;

import java.nio.file.Files;
import java.nio.file.Path;
import java.time.LocalDate;
import java.util.List;
import java.util.Map;
import java.util.zip.ZipFile;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;
import static org.mockito.Mockito.*;

class AiosGeneracionServiceTest {

    @TempDir
    Path tempDir;

    @Test
    void shouldGenerateAllMonthlyPeriodsInOneWorkbook() {
        MensualDataReader reader = mock(MensualDataReader.class);
        MensualExcelGenerator generator = mock(MensualExcelGenerator.class);
        AiosGeneracionService service = new AiosGeneracionService(
                reader,
                generator,
                mock(SemestralExcelGenerator.class),
                mock(TrimestralDataReader.class),
                mock(TrimestralExcelGenerator.class)
        );
        MensualData junio = mock(MensualData.class);
        MensualData julio = mock(MensualData.class);
        Path salida = Path.of("target/aios-output/Boletin_AIOS MENSUAL.xlsx");
        when(reader.read(LocalDate.of(2025, 6, 30))).thenReturn(junio);
        when(reader.read(LocalDate.of(2025, 7, 31))).thenReturn(julio);
        when(generator.generar(List.of(junio, julio))).thenReturn(salida);

        var resultado = service.generarMensuales(
                LocalDate.of(2025, 6, 1),
                LocalDate.of(2025, 7, 31)
        );

        assertEquals(List.of(salida), resultado.archivosGenerados());
        verify(reader).read(LocalDate.of(2025, 6, 30));
        verify(reader).read(LocalDate.of(2025, 7, 31));
        verify(generator).generar(List.of(junio, julio));
    }

    @Test
    void shouldRejectInvertedMonthlyRange() {
        AiosGeneracionService service = new AiosGeneracionService(
                mock(MensualDataReader.class),
                mock(MensualExcelGenerator.class),
                mock(SemestralExcelGenerator.class),
                mock(TrimestralDataReader.class),
                mock(TrimestralExcelGenerator.class)
        );

        IllegalArgumentException ex = assertThrows(IllegalArgumentException.class,
                () -> service.generarMensuales(LocalDate.of(2025, 7, 1), LocalDate.of(2025, 6, 30)));

        assertEquals("La fecha inicial no puede ser posterior a la fecha final", ex.getMessage());
    }

    @Test
    void shouldGenerateQuarterlyRangeInOneWorkbook() {
        TrimestralDataReader reader = mock(TrimestralDataReader.class);
        TrimestralExcelGenerator generator = mock(TrimestralExcelGenerator.class);
        AiosGeneracionService service = new AiosGeneracionService(
                mock(MensualDataReader.class),
                mock(MensualExcelGenerator.class),
                mock(SemestralExcelGenerator.class),
                reader,
                generator
        );
        LocalDate junio = LocalDate.of(2025, 6, 30);
        LocalDate septiembre = LocalDate.of(2025, 9, 30);
        LocalDate diciembre = LocalDate.of(2025, 12, 31);
        TrimestralData dataJunio = mock(TrimestralData.class);
        TrimestralData dataSeptiembre = mock(TrimestralData.class);
        TrimestralData dataDiciembre = mock(TrimestralData.class);
        when(reader.read(junio)).thenReturn(dataJunio);
        when(reader.read(septiembre)).thenReturn(dataSeptiembre);
        when(reader.read(diciembre)).thenReturn(dataDiciembre);
        List<TrimestralExcelGenerator.PeriodoTrimestral> periodos = List.of(
                new TrimestralExcelGenerator.PeriodoTrimestral(junio, dataJunio),
                new TrimestralExcelGenerator.PeriodoTrimestral(septiembre, dataSeptiembre),
                new TrimestralExcelGenerator.PeriodoTrimestral(diciembre, dataDiciembre)
        );
        Path salida = Path.of("target/aios-output/Boletin_AIOS TRIMESTRAL.xlsx");
        when(generator.generar(periodos)).thenReturn(salida);

        var resultado = service.generarRango(
                LocalDate.of(2025, 6, 1), LocalDate.of(2025, 12, 31), ModoGeneracion.TRIMESTRAL);

        assertEquals(List.of(salida), resultado.archivosGenerados());
        verify(generator).generar(periodos);
    }

    @Test
    void shouldGenerateSemesterRangeInOneWorkbook() {
        MensualDataReader mensualReader = mock(MensualDataReader.class);
        TrimestralDataReader trimestralReader = mock(TrimestralDataReader.class);
        SemestralExcelGenerator generator = mock(SemestralExcelGenerator.class);
        AiosGeneracionService service = new AiosGeneracionService(
                mensualReader,
                mock(MensualExcelGenerator.class),
                generator,
                trimestralReader,
                mock(TrimestralExcelGenerator.class)
        );
        LocalDate junio = LocalDate.of(2025, 6, 30);
        LocalDate diciembre = LocalDate.of(2025, 12, 31);
        MensualData mensualJunio = mock(MensualData.class);
        MensualData mensualDiciembre = mock(MensualData.class);
        TrimestralData trimestralJunio = mock(TrimestralData.class);
        TrimestralData trimestralDiciembre = mock(TrimestralData.class);
        when(mensualReader.read(junio)).thenReturn(mensualJunio);
        when(mensualReader.read(diciembre)).thenReturn(mensualDiciembre);
        when(trimestralReader.readForSemestral(junio, mensualJunio)).thenReturn(trimestralJunio);
        when(trimestralReader.readForSemestral(diciembre, mensualDiciembre)).thenReturn(trimestralDiciembre);
        List<SemestralExcelGenerator.PeriodoSemestral> periodos = List.of(
                new SemestralExcelGenerator.PeriodoSemestral(junio, mensualJunio, trimestralJunio),
                new SemestralExcelGenerator.PeriodoSemestral(diciembre, mensualDiciembre, trimestralDiciembre)
        );
        Path salida = Path.of("target/aios-output/semestral.xlsx");
        when(generator.generar(periodos)).thenReturn(salida);

        var resultado = service.generarRango(
                LocalDate.of(2025, 6, 1), LocalDate.of(2025, 12, 31), ModoGeneracion.SEMESTRAL);

        assertEquals(List.of(salida), resultado.archivosGenerados());
        verify(generator).generar(periodos);
    }

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
                one, one, one, one, one, one, one,
                one, one, one, one, one, one,
                one, one, one, one, one, one,
                one, one, one, one, one, one,
                one, one, one, one, one, one,
                one, one, one, one, one,
                one);
        TrimestralData data = new TrimestralData("jun-25", Map.of(), Map.of(), Map.of(), Map.of(), Map.of(), Map.of(), Map.of(), Map.of());
        when(mensualDataReader.read(fecha)).thenReturn(mensual);
        when(trimestralDataReader.readForSemestral(fecha, mensual)).thenReturn(data);
        when(semestralExcelGenerator.generar(fecha, mensual, data)).thenReturn(Path.of("target/aios-output/semestral.xlsx"));

        var resultado = service.generar(fecha, ModoGeneracion.SEMESTRAL);

        assertEquals(1, resultado.archivosGenerados().size());
        assertEquals("semestral.xlsx", resultado.archivosGenerados().getFirst().getFileName().toString());
        verify(mensualDataReader).read(fecha);
        verify(trimestralDataReader).readForSemestral(fecha, mensual);
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

    @Test
    void shouldContinueGeneratingTodoWhenSemestralFails() throws Exception {
        MensualDataReader mensualDataReader = mock(MensualDataReader.class);
        MensualExcelGenerator mensualExcelGenerator = mock(MensualExcelGenerator.class);
        SemestralExcelGenerator semestralExcelGenerator = mock(SemestralExcelGenerator.class);
        TrimestralDataReader trimestralDataReader = mock(TrimestralDataReader.class);
        TrimestralExcelGenerator trimestralExcelGenerator = mock(TrimestralExcelGenerator.class);
        AiosGeneracionService service = new AiosGeneracionService(
                mensualDataReader,
                mensualExcelGenerator,
                semestralExcelGenerator,
                trimestralDataReader,
                trimestralExcelGenerator,
                tempDir
        );

        LocalDate fecha = LocalDate.of(2025, 6, 30);
        MensualData mensual = mock(MensualData.class);
        TrimestralData trimestral = mock(TrimestralData.class);
        Path archivoMensual = Files.writeString(tempDir.resolve("mensual.xlsx"), "mensual");
        Path archivoTrimestral = Files.writeString(tempDir.resolve("trimestral.xlsx"), "trimestral");

        when(mensualDataReader.read(fecha)).thenReturn(mensual);
        when(mensualExcelGenerator.generar(mensual)).thenReturn(archivoMensual);
        when(trimestralDataReader.read(fecha)).thenReturn(trimestral);
        when(trimestralExcelGenerator.generar(fecha, trimestral)).thenReturn(archivoTrimestral);
        when(trimestralDataReader.readForSemestral(fecha, mensual)).thenReturn(trimestral);
        when(semestralExcelGenerator.generar(fecha, mensual, trimestral))
                .thenThrow(new IllegalStateException("No fue posible generar archivo semestral"));

        var resultado = service.generar(fecha, ModoGeneracion.TODO);

        assertTrue(resultado.zip());
        assertEquals(1, resultado.archivosGenerados().size());
        Path zip = resultado.archivosGenerados().getFirst();
        assertTrue(Files.exists(zip));
        try (ZipFile zipFile = new ZipFile(zip.toFile())) {
            assertTrue(zipFile.stream().anyMatch(entry -> entry.getName().equals("mensual.xlsx")));
            assertTrue(zipFile.stream().anyMatch(entry -> entry.getName().equals("trimestral.xlsx")));
            assertFalse(zipFile.stream().anyMatch(entry -> entry.getName().contains("semestral")));
        }
        verify(semestralExcelGenerator).generar(fecha, mensual, trimestral);
    }
}
