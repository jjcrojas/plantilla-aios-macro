package co.gov.sfc.excel;

import org.junit.jupiter.api.Test;
import org.springframework.jdbc.core.JdbcTemplate;

import java.math.BigDecimal;
import java.sql.Date;
import java.time.LocalDate;
import java.util.Map;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;
import static org.mockito.ArgumentMatchers.contains;
import static org.mockito.ArgumentMatchers.eq;
import static org.mockito.Mockito.mock;
import static org.mockito.Mockito.when;

class GastosTrimestralQueryServiceTest {

    @Test
    void shouldReadConsolidatedOperatingExpensesForEveryAfp() {
        JdbcTemplate jdbcTemplate = mock(JdbcTemplate.class);
        stubEntity(jdbcTemplate, 10, "10.25");
        stubEntity(jdbcTemplate, 3, "20.50");
        stubEntity(jdbcTemplate, 2, "30.75");
        stubEntity(jdbcTemplate, 9, "40.00");
        GastosTrimestralQueryService service = new GastosTrimestralQueryService(jdbcTemplate);

        Map<String, BigDecimal> values = service.leerGastosUsd(
                LocalDate.of(2025, 6, 30), new BigDecimal("4069.67"));

        assertEquals(new BigDecimal("10.25"), values.get("colf"));
        assertEquals(new BigDecimal("20.50"), values.get("porv"));
        assertEquals(new BigDecimal("30.75"), values.get("prot"));
        assertEquals(new BigDecimal("40.00"), values.get("sk"));
        String sql = service.sqlGastosPorEntidad();
        assertTrue(sql.contains("p.Codigo = 510000"));
        assertTrue(sql.contains("p.Codigo IN (510300, 510400, 510600, 510700, 510800"));
        assertTrue(sql.contains("e.Codigo_Entidad = ?"));
        assertTrue(sql.contains("ei.Tipo_Informe = 0"));
    }

    private void stubEntity(JdbcTemplate jdbcTemplate, int codigoEntidad, String result) {
        Date fechaCorte = Date.valueOf("2025-06-30");
        Date cierreAnterior = Date.valueOf("2024-12-31");
        Date corteAnterior = Date.valueOf("2024-06-30");
        when(jdbcTemplate.queryForObject(contains("e.Codigo_Entidad = ?"), eq(BigDecimal.class),
                eq(fechaCorte), eq(cierreAnterior), eq(corteAnterior),
                eq(fechaCorte), eq(cierreAnterior), eq(corteAnterior),
                eq(new BigDecimal("4069.67")), eq(codigoEntidad),
                eq(fechaCorte), eq(cierreAnterior), eq(corteAnterior)))
                .thenReturn(new BigDecimal(result));
    }
}
