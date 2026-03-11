package co.gov.sfc.insumos;

import co.gov.sfc.config.AiosProperties;
import org.springframework.context.annotation.Primary;
import org.springframework.stereotype.Component;

import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;

@Component
@Primary
public class FileSystemInsumosProvider implements InsumosProvider {

    private final AiosProperties properties;

    public FileSystemInsumosProvider(AiosProperties properties) {
        this.properties = properties;
    }

    @Override
    public InputStream open(String fileName) throws IOException {
        return Files.newInputStream(resolve(fileName));
    }

    @Override
    public Path resolve(String fileName) {
        return properties.insumosDir().resolve(fileName);
    }
}
