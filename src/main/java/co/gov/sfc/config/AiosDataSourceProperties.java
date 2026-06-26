package co.gov.sfc.config;

import org.springframework.boot.context.properties.ConfigurationProperties;

@ConfigurationProperties(prefix = "aios.datasource")
public record AiosDataSourceProperties(
        String url,
        String driverClassName,
        String username,
        String password,
        int maximumPoolSize,
        int minimumIdle,
        long connectionTimeout,
        long validationTimeout,
        long initializationFailTimeout
) {
}
