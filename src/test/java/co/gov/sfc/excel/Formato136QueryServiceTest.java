package co.gov.sfc.excel;

import org.junit.jupiter.api.Test;
import org.springframework.jdbc.core.JdbcTemplate;

import java.math.BigDecimal;
import java.sql.Date;
import java.time.LocalDate;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.mockito.ArgumentMatchers.eq;
import static org.mockito.ArgumentMatchers.contains;
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
}
