package co.gov.sfc.excel;

import org.junit.jupiter.api.Test;
import org.springframework.jdbc.core.JdbcTemplate;
import org.springframework.jdbc.core.RowCallbackHandler;

import java.math.BigDecimal;
import java.sql.Date;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.time.LocalDate;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.mockito.ArgumentMatchers.any;
import static org.mockito.ArgumentMatchers.contains;
import static org.mockito.ArgumentMatchers.eq;
import static org.mockito.Mockito.doAnswer;
import static org.mockito.Mockito.mock;
import static org.mockito.Mockito.verify;
import static org.mockito.Mockito.when;

class Formato136QueryServiceTest {

    @Test
    void shouldReadAportesRecibidosFromTeradataForCutoffWindow() {
        JdbcTemplate jdbcTemplate = mock(JdbcTemplate.class);
        when(jdbcTemplate.queryForObject(contains("d.nivel1 = 136"), eq(BigDecimal.class),
                eq(Date.valueOf("2024-06-01")), eq(Date.valueOf("2025-06-30"))))
                .thenReturn(new BigDecimal("123.45"));

        Formato136QueryService service = new Formato136QueryService(jdbcTemplate);
        BigDecimal value = service.leerAportesRecibidos(LocalDate.of(2025, 6, 30));

        assertEquals(new BigDecimal("123.45"), value);
        verify(jdbcTemplate).queryForObject(contains("SUM(e.valor) / 1000000"), eq(BigDecimal.class),
                eq(Date.valueOf("2024-06-01")), eq(Date.valueOf("2025-06-30")));
    }

    @Test
    void shouldReadColombiaTrimestralFromTeradataByFundAndEntity() throws SQLException {
        JdbcTemplate jdbcTemplate = mock(JdbcTemplate.class);
        doAnswer(invocation -> {
            RowCallbackHandler handler = invocation.getArgument(1);
            handler.processRow(row(10, 1000, "100.00"));
            handler.processRow(row(9, 8000, "25.00"));
            return null;
        }).when(jdbcTemplate).query(contains("GROUP BY a.codigo_entidad, c.codigo_patrimonio"),
                any(RowCallbackHandler.class), eq(Date.valueOf("2025-06-30")));

        Formato136QueryService service = new Formato136QueryService(jdbcTemplate);
        var valores = service.leerColombiaPorFondoEntidad(LocalDate.of(2025, 6, 30));

        assertEquals(new BigDecimal("100.00"), valores.get("mod_colf"));
        assertEquals(new BigDecimal("25.00"), valores.get("mod_alt"));
    }

    private ResultSet row(int codigoEntidad, int codigoPatrimonio, String valor) throws SQLException {
        ResultSet rs = mock(ResultSet.class);
        when(rs.getInt("codigo_entidad")).thenReturn(codigoEntidad);
        when(rs.getInt("codigo_patrimonio")).thenReturn(codigoPatrimonio);
        when(rs.getBigDecimal("valor_total")).thenReturn(new BigDecimal(valor));
        return rs;
    }
}
