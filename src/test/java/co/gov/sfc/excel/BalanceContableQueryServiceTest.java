package co.gov.sfc.excel;

import org.junit.jupiter.api.Test;
import org.springframework.jdbc.core.JdbcTemplate;

import java.math.BigDecimal;
import java.sql.Date;
import java.time.LocalDate;
import java.util.List;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;
import static org.mockito.ArgumentMatchers.anyString;
import static org.mockito.ArgumentMatchers.eq;
import static org.mockito.Mockito.mock;
import static org.mockito.Mockito.verify;
import static org.mockito.Mockito.when;

class BalanceContableQueryServiceTest {

    @Test
    void shouldReadActivoPasivoPatrimonioAndActiveAdministratorsWithoutQuotes() {
        JdbcTemplate jdbcTemplate = mock(JdbcTemplate.class);
        Date fecha = Date.valueOf("2025-06-30");
        when(jdbcTemplate.queryForList(anyString(), eq(String.class)))
                .thenReturn(List.of("\"PORVENIR\"", "“PROTECCIÓN”", "'SKANDIA'", "PORVENIR"));
        when(jdbcTemplate.queryForObject(anyString(), eq(BigDecimal.class), eq(100000), eq(fecha)))
                .thenReturn(new BigDecimal("987654.32"));
        when(jdbcTemplate.queryForObject(anyString(), eq(BigDecimal.class), eq(200000), eq(fecha)))
                .thenReturn(new BigDecimal("123456.78"));
        when(jdbcTemplate.queryForObject(anyString(), eq(BigDecimal.class), eq(300000), eq(fecha)))
                .thenReturn(new BigDecimal("864197.54"));
        BalanceContableQueryService service = new BalanceContableQueryService(jdbcTemplate);

        BalanceContableQueryService.BalanceContable balance =
                service.leer(LocalDate.of(2025, 6, 30));

        assertEquals(new BigDecimal("987654.32"), balance.activoMmCop());
        assertEquals(new BigDecimal("123456.78"), balance.pasivoMmCop());
        assertEquals(new BigDecimal("864197.54"), balance.patrimonioMmCop());
        assertEquals(List.of("PORVENIR", "PROTECCIÓN", "SKANDIA"), balance.administradorasVigentes());
        assertEquals(new BigDecimal("3"), balance.numeroAdministradorasVigentes());
        verify(jdbcTemplate).queryForObject(anyString(), eq(BigDecimal.class), eq(100000), eq(fecha));
        verify(jdbcTemplate).queryForObject(anyString(), eq(BigDecimal.class), eq(200000), eq(fecha));
        verify(jdbcTemplate).queryForObject(anyString(), eq(BigDecimal.class), eq(300000), eq(fecha));

        String sqlCuenta = service.sqlCuenta();
        assertTrue(sqlCuenta.contains("PROD_DWH_CONSULTA.ESTFIN_INDIV"));
        assertTrue(sqlCuenta.contains("e.Tipo_Entidad = 23"));
        assertTrue(sqlCuenta.contains("ef.Tipo_Informe = 0"));
        assertTrue(sqlCuenta.contains("p.Codigo = ?"));
        assertTrue(sqlCuenta.contains("t.Fecha = ?"));
        assertFalse(sqlCuenta.contains("e.Estado"));
        assertTrue(service.sqlAdministradorasVigentes().contains("e.Estado = 1"));
    }
}
