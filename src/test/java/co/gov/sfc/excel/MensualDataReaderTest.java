package co.gov.sfc.excel;

import co.gov.sfc.config.AiosProperties;
import co.gov.sfc.insumos.InsumosLocator;
import org.junit.jupiter.api.Test;

import java.math.BigDecimal;
import java.nio.file.Path;
import java.time.LocalDate;
import java.util.List;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.mockito.Mockito.mock;
import static org.mockito.Mockito.verify;
import static org.mockito.Mockito.when;

class MensualDataReaderTest {

    @Test
    void shouldUseSameOneYearRentabilidadServiceAsSemestral() {
        LocalDate fechaCorte = LocalDate.of(2025, 6, 30);
        LocalDate fechaInicio = LocalDate.of(2024, 6, 30);
        Path rentFile = Path.of("target", "archivo-no-necesita-existir", "Rent_Vr_Uni_Moderado.xlsm");

        InsumosLocator locator = mock(InsumosLocator.class);
        AiosProperties properties = mock(AiosProperties.class);
        Formato491QueryService formato491 = mock(Formato491QueryService.class);
        FondoAdministradoQueryService fondoAdministrado = mock(FondoAdministradoQueryService.class);
        Formato493QueryService formato493 = mock(Formato493QueryService.class);
        Formato495QueryService formato495 = mock(Formato495QueryService.class);
        TrmService trmService = mock(TrmService.class);
        SeriesEconomicasService seriesEconomicas = mock(SeriesEconomicasService.class);
        BalanceContableQueryService balanceContable = mock(BalanceContableQueryService.class);
        RentabilidadService rentabilidadService = mock(RentabilidadService.class);

        when(locator.findRequired("Rent_Vr_Uni_Moderado", fechaCorte)).thenReturn(rentFile);
        when(formato491.leerResumen(fechaCorte)).thenReturn(resumen491EnCeros());
        when(formato493.leerTraspasosSistema(fechaCorte)).thenReturn(BigDecimal.ZERO);
        when(fondoAdministrado.leer(fechaCorte)).thenReturn(
                new FondoAdministradoQueryService.FondoAdministrado(BigDecimal.ZERO, BigDecimal.ZERO));
        when(trmService.obtener(fechaCorte)).thenReturn(new BigDecimal("4069.67"));
        when(seriesEconomicas.leer(fechaCorte)).thenReturn(
                new SeriesEconomicasService.SeriesEconomicas(
                        BigDecimal.ZERO, BigDecimal.ZERO, BigDecimal.ZERO, Path.of("series.xlsx")));
        when(balanceContable.leer(fechaCorte)).thenReturn(
                new BalanceContableQueryService.BalanceContable(
                        BigDecimal.ZERO, BigDecimal.ZERO, BigDecimal.ZERO, List.of()));
        when(formato495.leerResumen(fechaCorte)).thenReturn(
                new Formato495QueryService.PensionadosResumen(
                        BigDecimal.ZERO, BigDecimal.ZERO, BigDecimal.ZERO, BigDecimal.ZERO));
        when(rentabilidadService.calcularRentabilidad(rentFile, fechaCorte, 1)).thenReturn(
                new RentabilidadService.RentabilidadResultado(
                        fechaInicio,
                        fechaCorte,
                        new BigDecimal("0.1022608588509073"),
                        new BigDecimal("0.051511390166620874")));

        MensualDataReader reader = new MensualDataReader(
                locator,
                properties,
                formato491,
                fondoAdministrado,
                formato493,
                formato495,
                trmService,
                seriesEconomicas,
                balanceContable,
                rentabilidadService);

        MensualData result = reader.read(fechaCorte);

        assertEquals(new BigDecimal("0.1022608588509073"), result.tmpNominal1());
        assertEquals(new BigDecimal("0.051511390166620874"), result.tmpReal1());
        verify(rentabilidadService).calcularRentabilidad(rentFile, fechaCorte, 1);
    }

    private Formato491QueryService.Resumen491 resumen491EnCeros() {
        return new Formato491QueryService.Resumen491(
                BigDecimal.ZERO,
                BigDecimal.ZERO,
                BigDecimal.ZERO,
                BigDecimal.ZERO,
                BigDecimal.ZERO,
                BigDecimal.ZERO,
                BigDecimal.ZERO,
                BigDecimal.ZERO,
                BigDecimal.ZERO,
                BigDecimal.ZERO,
                BigDecimal.ZERO);
    }
}
