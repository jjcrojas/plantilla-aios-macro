package co.gov.sfc.config;

import com.zaxxer.hikari.HikariDataSource;
import org.junit.jupiter.api.Test;

import static org.junit.jupiter.api.Assertions.assertEquals;

class AiosDataSourceConfigurationTest {

    @Test
    void shouldBuildAiosDataSourceFromDedicatedProperties() {
        AiosDataSourceProperties properties = new AiosDataSourceProperties(
                "jdbc:teradata://10.40.176.8/DATABASE=prod_dwh_consulta,LOGMECH=LDAP",
                "com.teradata.jdbc.TeraDriver",
                "aios-user",
                "aios-password",
                4,
                1,
                30_000,
                5_000,
                -1
        );

        try (HikariDataSource dataSource = (HikariDataSource) new AiosDataSourceConfiguration().dataSource(properties)) {
            assertEquals(properties.url(), dataSource.getJdbcUrl());
            assertEquals(properties.username(), dataSource.getUsername());
            assertEquals(properties.password(), dataSource.getPassword());
            assertEquals("AiosTeradataPool", dataSource.getPoolName());
        }
    }
}
