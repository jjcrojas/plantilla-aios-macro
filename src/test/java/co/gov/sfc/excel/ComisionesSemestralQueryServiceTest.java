package co.gov.sfc.excel;

import org.junit.jupiter.api.Test;
import org.mockito.ArgumentCaptor;
import org.springframework.jdbc.core.JdbcTemplate;

import java.math.BigDecimal;
import java.sql.Date;
import java.time.LocalDate;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;
import static org.mockito.ArgumentMatchers.eq;
import static org.mockito.Mockito.mock;
import static org.mockito.Mockito.verify;
import static org.mockito.Mockito.when;

class ComisionesSemestralQueryServiceTest {

    @Test
    void shouldCalculateAccount411500ForJuneUsingThreeDatesAndTrm() {
        JdbcTemplate jdbcTemplate = mock(JdbcTemplate.class);
        when(jdbcTemplate.queryForObject(
                org.mockito.ArgumentMatchers.anyString(), eq(BigDecimal.class),
                eq(Date.valueOf("2025-06-30")), eq(Date.valueOf("2024-12-31")), eq(Date.valueOf("2024-06-30")),
                eq(new BigDecimal("4069.67")),
                eq(Date.valueOf("2025-06-30")), eq(Date.valueOf("2024-12-31")), eq(Date.valueOf("2024-06-30")),
                eq(411500)))
                .thenReturn(new BigDecimal("123.456"));
        ComisionesSemestralQueryService service = new ComisionesSemestralQueryService(jdbcTemplate);

        BigDecimal result = service.leer411500(LocalDate.of(2025, 6, 30), new BigDecimal("4069.67"));

        assertEquals(new BigDecimal("123.456"), result);
        ArgumentCaptor<String> sql = ArgumentCaptor.forClass(String.class);
        verify(jdbcTemplate).queryForObject(
                sql.capture(), eq(BigDecimal.class),
                eq(Date.valueOf("2025-06-30")), eq(Date.valueOf("2024-12-31")), eq(Date.valueOf("2024-06-30")),
                eq(new BigDecimal("4069.67")),
                eq(Date.valueOf("2025-06-30")), eq(Date.valueOf("2024-12-31")), eq(Date.valueOf("2024-06-30")),
                eq(411500));
        assertTrue(sql.getValue().contains("PROD_DWH_CONSULTA.ESTFIN_INDIV"));
        assertTrue(sql.getValue().contains("ei.Tipo_Informe = 0"));
        assertTrue(sql.getValue().contains("p.Codigo = ?"));
        assertTrue(sql.getValue().contains("/ 1000000 / ?"));
    }

    @Test
    void shouldCalculateAnyRequestedAccountUsingTheSameSemesterFlowFormula() {
        JdbcTemplate jdbcTemplate = mock(JdbcTemplate.class);
        when(jdbcTemplate.queryForObject(
                org.mockito.ArgumentMatchers.anyString(), eq(BigDecimal.class),
                eq(Date.valueOf("2025-06-30")), eq(Date.valueOf("2024-12-31")), eq(Date.valueOf("2024-06-30")),
                eq(new BigDecimal("4069.67")),
                eq(Date.valueOf("2025-06-30")), eq(Date.valueOf("2024-12-31")), eq(Date.valueOf("2024-06-30")),
                eq(519015)))
                .thenReturn(new BigDecimal("98.765"));
        ComisionesSemestralQueryService service = new ComisionesSemestralQueryService(jdbcTemplate);

        BigDecimal result = service.leerCuenta(
                LocalDate.of(2025, 6, 30), new BigDecimal("4069.67"), 519015);

        assertEquals(new BigDecimal("98.765"), result);
        String sql = service.sqlCuenta();
        assertTrue(sql.contains("p.Codigo = ?"));
        assertTrue(sql.contains("t.Fecha IN (?, ?, ?)"));
        assertTrue(sql.contains("/ 1000000 / ?"));
    }

    @Test
    void shouldCalculateSemestralRow54FromAccount590000() {
        JdbcTemplate jdbcTemplate = mock(JdbcTemplate.class);
        when(jdbcTemplate.queryForObject(
                org.mockito.ArgumentMatchers.anyString(), eq(BigDecimal.class),
                eq(Date.valueOf("2025-06-30")), eq(Date.valueOf("2024-12-31")), eq(Date.valueOf("2024-06-30")),
                eq(new BigDecimal("4069.67")),
                eq(Date.valueOf("2025-06-30")), eq(Date.valueOf("2024-12-31")), eq(Date.valueOf("2024-06-30")),
                eq(590000)))
                .thenReturn(new BigDecimal("147.25"));
        ComisionesSemestralQueryService service = new ComisionesSemestralQueryService(jdbcTemplate);

        BigDecimal result = service.leer590000(
                LocalDate.of(2025, 6, 30), new BigDecimal("4069.67"));

        assertEquals(new BigDecimal("147.25"), result);
        verify(jdbcTemplate).queryForObject(
                org.mockito.ArgumentMatchers.anyString(), eq(BigDecimal.class),
                eq(Date.valueOf("2025-06-30")), eq(Date.valueOf("2024-12-31")), eq(Date.valueOf("2024-06-30")),
                eq(new BigDecimal("4069.67")),
                eq(Date.valueOf("2025-06-30")), eq(Date.valueOf("2024-12-31")), eq(Date.valueOf("2024-06-30")),
                eq(590000));
    }

    @Test
    void shouldCalculateOperatingExpensesForJuneUsingSignedAccountsDatesAndTrm() {
        JdbcTemplate jdbcTemplate = mock(JdbcTemplate.class);
        when(jdbcTemplate.queryForObject(
                org.mockito.ArgumentMatchers.anyString(), eq(BigDecimal.class),
                eq(Date.valueOf("2025-06-30")), eq(Date.valueOf("2024-12-31")), eq(Date.valueOf("2024-06-30")),
                eq(Date.valueOf("2025-06-30")), eq(Date.valueOf("2024-12-31")), eq(Date.valueOf("2024-06-30")),
                eq(new BigDecimal("4069.67")),
                eq(Date.valueOf("2025-06-30")), eq(Date.valueOf("2024-12-31")), eq(Date.valueOf("2024-06-30"))))
                .thenReturn(new BigDecimal("456.789"));
        ComisionesSemestralQueryService service = new ComisionesSemestralQueryService(jdbcTemplate);

        BigDecimal result = service.leerGastosOperativos(LocalDate.of(2025, 6, 30), new BigDecimal("4069.67"));

        assertEquals(new BigDecimal("456.789"), result);
        String sql = service.sqlGastosOperativos();
        assertTrue(sql.contains("p.Codigo = 510000"));
        assertTrue(sql.contains("510300, 510400, 510600, 510700, 510800"));
        assertTrue(sql.contains("512500, 512800, 512900, 513900"));
        assertTrue(sql.contains("THEN -("));
        assertTrue(sql.contains("ei.Tipo_Informe = 0"));
        assertTrue(sql.contains("/ 1000000 / ?"));
    }
}
