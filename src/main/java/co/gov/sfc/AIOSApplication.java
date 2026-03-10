package co.gov.sfc;

import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.boot.context.properties.ConfigurationPropertiesScan;

@SpringBootApplication
@ConfigurationPropertiesScan
public class AIOSApplication {

    public static void main(String[] args) {
        SpringApplication.run(AIOSApplication.class, args);
    }
}
