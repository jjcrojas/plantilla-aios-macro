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
            handler.processRow(row(10, "ENT_10", "7890000.00"));
            return null;
        }).when(jdbcTemplate).query(contains("pa.Codigo_Patrimonio = ?"),
                any(RowCallbackHandler.class), eq(5000), eq(Date.valueOf("2025-06-30")));

        doAnswer(invocation -> {
            RowCallbackHandler handler = invocation.getArgument(1);
            handler.processRow(row(2, "ENT_2", "94247215063"));
            handler.processRow(row(3, "ENT_3", "128081241326"));
            handler.processRow(row(9, "ENT_9", "11964471789"));
            handler.processRow(row(9, "SKANDIA_ALT", "453166850"));
            handler.processRow(row(10, "ENT_10", "32491777302"));
            return null;
        }).when(jdbcTemplate).query(contains("pa.Codigo_Patrimonio IN (4, 8000)"),
                any(RowCallbackHandler.class), eq(Date.valueOf("2025-06-30")));

        doAnswer(invocation -> {
            RowCallbackHandler handler = invocation.getArgument(1);
            handler.processRow(row(3, "ENT_3", "1200000.00"));
            return null;
        }).when(jdbcTemplate).query(contains("pa.Codigo_Patrimonio = ?"),
                any(RowCallbackHandler.class), eq(7000), eq(Date.valueOf("2025-06-30")));

        doAnswer(invocation -> {
            RowCallbackHandler handler = invocation.getArgument(1);
            handler.processRow(row(9, "ENT_9", "800000.00"));
            return null;
        }).when(jdbcTemplate).query(contains("pa.Codigo_Patrimonio = ?"),
                any(RowCallbackHandler.class), eq(6000), eq(Date.valueOf("2025-06-30")));

        Formato136QueryService service = new Formato136QueryService(jdbcTemplate);
        var valores = service.leerColombiaPorFondoEntidad(LocalDate.of(2025, 6, 30));

        assertNumericEquals("7890", valores.get("con_colf"));
        assertNumericEquals("94247215.063", valores.get("mod_prot"));
        assertNumericEquals("128081241.326", valores.get("mod_porv"));
        assertNumericEquals("11964471.789", valores.get("mod_sk"));
        assertNumericEquals("453166.85", valores.get("mod_alt"));
        assertNumericEquals("32491777.302", valores.get("mod_colf"));
        assertNumericEquals("1200", valores.get("rp_porv"));
        assertNumericEquals("800", valores.get("mr_sk"));

        verify(jdbcTemplate).query(contains("FROM PROD_DWH_CONSULTA.ESTFIN_INDIV_PA"),
                any(RowCallbackHandler.class), eq(5000), eq(Date.valueOf("2025-06-30")));
        verify(jdbcTemplate).query(contains("p.Codigo = 100000"),
                any(RowCallbackHandler.class), eq(Date.valueOf("2025-06-30")));
    }

    private void assertNumericEquals(String expected, BigDecimal actual) {
        assertEquals(0, new BigDecimal(expected).compareTo(actual));
    }

    private ResultSet row(int codigoEntidad, String claveReporte, String valorMiles) throws SQLException {
        ResultSet rs = mock(ResultSet.class);
        when(rs.getInt("codigo_entidad")).thenReturn(codigoEntidad);
        when(rs.getString("clave_reporte")).thenReturn(claveReporte);
        when(rs.getBigDecimal("valor_miles")).thenReturn(new BigDecimal(valorMiles));
        return rs;
    }
}
