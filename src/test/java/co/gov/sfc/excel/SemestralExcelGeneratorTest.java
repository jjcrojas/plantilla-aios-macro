package co.gov.sfc.excel;

import co.gov.sfc.config.AiosProperties;
import co.gov.sfc.insumos.InsumosLocator;
import org.junit.jupiter.api.Test;

import java.lang.reflect.Method;
import java.math.BigDecimal;
import java.nio.file.Path;
import java.time.LocalDate;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.mockito.Mockito.mock;
import static org.mockito.Mockito.when;

class SemestralExcelGeneratorTest {

    @Test
    void shouldReadFila25FromFallecidosSheetForJune2025() throws Exception {
        AiosProperties properties = new AiosProperties(Path.of("insumos_ejemplo"), null, null, null, null);
        Formato493QueryService formato493QueryService = mock(Formato493QueryService.class);
        when(formato493QueryService.leerFallecidosSistema(LocalDate.of(2025, 6, 30))).thenReturn(new BigDecimal("38279"));
        SemestralExcelGenerator generator = new SemestralExcelGenerator(properties, new InsumosLocator(properties), null, formato493QueryService, mock(Formato495QueryService.class), mock(Formato136QueryService.class));
        Method method = SemestralExcelGenerator.class.getDeclaredMethod("readFila25Trimestral493", LocalDate.class);
        method.setAccessible(true);

        BigDecimal value = (BigDecimal) method.invoke(generator, LocalDate.of(2025, 6, 30));

        assertEquals(new BigDecimal("38.27900000"), value);
    }
}
