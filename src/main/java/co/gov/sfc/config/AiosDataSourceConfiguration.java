package co.gov.sfc.config;

import com.zaxxer.hikari.HikariConfig;
import com.zaxxer.hikari.HikariDataSource;
import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Configuration;

import javax.sql.DataSource;

@Configuration(proxyBeanMethods = false)
public class AiosDataSourceConfiguration {

    @Bean
    public DataSource dataSource(AiosDataSourceProperties properties) {
        HikariConfig config = new HikariConfig();
        config.setPoolName("AiosTeradataPool");
        config.setJdbcUrl(properties.url());
        config.setDriverClassName(properties.driverClassName());
        config.setUsername(properties.username());
        config.setPassword(properties.password());
        config.setMaximumPoolSize(properties.maximumPoolSize());
        config.setMinimumIdle(properties.minimumIdle());
        config.setConnectionTimeout(properties.connectionTimeout());
        config.setValidationTimeout(properties.validationTimeout());
        config.setInitializationFailTimeout(properties.initializationFailTimeout());
        return new HikariDataSource(config);
    }
}
