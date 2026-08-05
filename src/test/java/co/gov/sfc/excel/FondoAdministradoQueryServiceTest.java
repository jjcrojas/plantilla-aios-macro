package co.gov.sfc.excel;

import org.junit.jupiter.api.Test;
import org.mockito.ArgumentCaptor;
import org.springframework.jdbc.core.JdbcTemplate;
import org.springframework.jdbc.core.RowMapper;

import java.math.BigDecimal;
import java.sql.Date;
import java.time.LocalDate;
import java.util.List;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;
import static org.mockito.ArgumentMatchers.any;
import static org.mockito.ArgumentMatchers.anyString;
import static org.mockito.Mockito.mock;
import static org.mockito.Mockito.verify;
import static org.mockito.Mockito.when;

class FondoAdministradoQueryServiceTest {

    @SuppressWarnings("unchecked")
    @Test
    void shouldConvertBalancesToMillionsAndCalculateProtectionPorvenirConcentration() throws Exception {
        JdbcTemplate jdbcTemplate = mock(JdbcTemplate.class);
        when(jdbcTemplate.query(anyString(), any(RowMapper.class), any(Object[].class)))
                .thenAnswer(invocation -> {
                    RowMapper<Object> mapper = invocation.getArgument(1);
                    var rsProteccion = mock(java.sql.ResultSet.class);
                    when(rsProteccion.getInt("CODIGO_ENTIDAD")).thenReturn(2);
                    when(rsProteccion.getBigDecimal("VALOR_MM_COP")).thenReturn(new BigDecimal("300000"));
                    var rsPorvenir = mock(java.sql.ResultSet.class);
                    when(rsPorvenir.getInt("CODIGO_ENTIDAD")).thenReturn(3);
                    when(rsPorvenir.getBigDecimal("VALOR_MM_COP")).thenReturn(new BigDecimal("400000"));
                    var rsOtra = mock(java.sql.ResultSet.class);
                    when(rsOtra.getInt("CODIGO_ENTIDAD")).thenReturn(10);
                    when(rsOtra.getBigDecimal("VALOR_MM_COP")).thenReturn(new BigDecimal("300000"));
                    return List.of(mapper.mapRow(rsProteccion, 0), mapper.mapRow(rsPorvenir, 1), mapper.mapRow(rsOtra, 2));
                });

        var result = new FondoAdministradoQueryService(jdbcTemplate).leer(LocalDate.of(2025, 6, 30));

        assertEquals(0, new BigDecimal("1000000").compareTo(result.totalMmCop()));
        assertEquals(0, new BigDecimal("70.0000000000").compareTo(result.concentracionProteccionPorvenirPct()));

        ArgumentCaptor<String> sql = ArgumentCaptor.forClass(String.class);
        verify(jdbcTemplate).query(sql.capture(), any(RowMapper.class),
                org.mockito.ArgumentMatchers.eq(Date.valueOf("2025-06-30")));
        assertTrue(sql.getValue().contains("Saldo_Sincierre_Total_Moneda_0"));
        assertTrue(sql.getValue().contains("/ 1000000"));
        assertTrue(sql.getValue().contains("t.Fecha = ?"));
    }
}
