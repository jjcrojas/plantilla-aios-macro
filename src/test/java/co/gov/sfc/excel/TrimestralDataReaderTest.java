package co.gov.sfc.excel;

import co.gov.sfc.insumos.InsumosLocator;
import org.junit.jupiter.api.Disabled;
import org.junit.jupiter.api.Test;

import java.lang.reflect.Method;
import java.math.BigDecimal;
import java.nio.file.Path;
import java.time.LocalDate;
import java.util.Map;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;
import static org.mockito.Mockito.mock;
import static org.mockito.Mockito.when;

class TrimestralDataReaderTest {

    @Disabled("Integra con workbook 493 pesado; se valida en pruebas de servicio")
    @Test
    void shouldReadTraspasosPerAfpFrom493UsingMacroCodes() {
        MensualDataReader mensualDataReader = mock(MensualDataReader.class);
        InsumosLocator locator = mock(InsumosLocator.class);
        Formato491QueryService formato491QueryService = mock(Formato491QueryService.class);
        Formato493QueryService formato493QueryService = mock(Formato493QueryService.class);
        Formato136QueryService formato136QueryService = mock(Formato136QueryService.class);
        ComisionesSfcService comisionesSfcService = mock(ComisionesSfcService.class);

        LocalDate fecha = LocalDate.of(2025, 6, 30);
        when(locator.findRequired("493", fecha)).thenReturn(Path.of("insumos_ejemplo", "Serie_Formato_493 MOVIMIENTO AFILIADOS.xlsx"));
        when(locator.findRequired("SISTEMA TOTAL", fecha)).thenReturn(Path.of("insumos_ejemplo", "SISTEMA TOTAL Junio 2025.xls"));
        when(locator.findRequired("Rent_Vr_Uni_Moderado", fecha)).thenReturn(Path.of("insumos_ejemplo", "Rent_Vr_Uni_Moderado.xlsm"));
        when(locator.findRequired("comision fpo desde 2003", fecha)).thenReturn(Path.of("insumos_ejemplo", "comisión FPO desde 2003.xls"));
        when(mensualDataReader.read(fecha)).thenReturn(new MensualData(
                "jun-25",
                BigDecimal.ZERO, BigDecimal.ZERO, BigDecimal.ZERO, BigDecimal.ZERO, BigDecimal.ZERO, BigDecimal.ZERO,
                BigDecimal.ZERO, BigDecimal.ZERO, BigDecimal.ZERO, BigDecimal.ZERO, BigDecimal.ZERO,
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
                BigDecimal.ZERO,
                BigDecimal.ZERO,
                BigDecimal.ZERO
        ));

        when(formato491QueryService.leerAportantesPorEntidad(fecha)).thenReturn(Map.of());
        when(formato493QueryService.leerTraspasosPorEntidad(fecha)).thenReturn(Map.of(
                "colf", BigDecimal.ZERO, "porv", BigDecimal.ZERO, "prot", BigDecimal.ZERO, "sk", BigDecimal.ZERO));
        when(formato136QueryService.leerColombiaPorFondoEntidad(fecha)).thenReturn(Map.of());

        TrimestralDataReader reader = new TrimestralDataReader(mensualDataReader, locator, formato491QueryService, formato493QueryService, formato136QueryService, comisionesSfcService);
        TrimestralData data = reader.read(fecha);

        assertTrue(data.traspasos().getOrDefault("colf", BigDecimal.ZERO).signum() >= 0);
        assertTrue(data.traspasos().getOrDefault("porv", BigDecimal.ZERO).signum() >= 0);
        assertTrue(data.traspasos().getOrDefault("prot", BigDecimal.ZERO).signum() >= 0);
        assertTrue(data.traspasos().getOrDefault("sk", BigDecimal.ZERO).signum() >= 0);
    }

    @Test
    void shouldUseRequiredOcrCommissionsForSemestralRow71() throws Exception {
        ComisionesSfcService comisionesSfcService = mock(ComisionesSfcService.class);
        LocalDate fecha = LocalDate.of(2025, 6, 30);
        when(comisionesSfcService.leer(fecha)).thenReturn(Map.of(
                "col_obl", new BigDecimal("0.97"), "col_seg", new BigDecimal("2.03"),
                "por_obl", new BigDecimal("0.47"), "por_seg", new BigDecimal("2.53"),
                "pro_obl", new BigDecimal("0.47"), "pro_seg", new BigDecimal("2.53"),
                "ska_obl", new BigDecimal("2.05"), "ska_seg", new BigDecimal("0.95")));
        TrimestralDataReader reader = new TrimestralDataReader(null, null, null, null, null, comisionesSfcService);
        Method method = TrimestralDataReader.class.getDeclaredMethod("readComisionesOcrRequerido", LocalDate.class);
        method.setAccessible(true);

        @SuppressWarnings("unchecked")
        Map<String, BigDecimal> values = (Map<String, BigDecimal>) method.invoke(reader, fecha);

        assertEquals(Map.of(
                "col_obl", new BigDecimal("0.97"),
                "por_obl", new BigDecimal("0.47"),
                "pro_obl", new BigDecimal("0.47"),
                "ska_obl", new BigDecimal("2.05")), values);
    }}
