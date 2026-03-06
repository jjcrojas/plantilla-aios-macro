package co.gov.sfc.insumos;

import co.gov.sfc.config.AiosProperties;
import org.springframework.stereotype.Component;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Comparator;

@Component
public class InsumosLocator {

    private final AiosProperties properties;

    public InsumosLocator(AiosProperties properties) {
        this.properties = properties;
    }

    public Path findRequired(String contains) {
        try (var stream = Files.list(properties.insumosDir())) {
            return stream
                    .filter(Files::isRegularFile)
                    .filter(p -> p.getFileName().toString().toLowerCase().contains(contains.toLowerCase()))
                    .sorted(Comparator.comparing(Path::toString))
                    .findFirst()
                    .orElseThrow(() -> new IllegalArgumentException("No se encontró insumo que contenga: " + contains));
        } catch (IOException e) {
            throw new IllegalStateException("No fue posible leer directorio de insumos", e);
        }
    }
}
