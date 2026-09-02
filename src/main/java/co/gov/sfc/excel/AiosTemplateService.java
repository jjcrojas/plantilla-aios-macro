package co.gov.sfc.excel;

import co.gov.sfc.config.AiosProperties;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Arrays;

/**
 * Abre las plantillas de presentación sin depender de datos históricos ni de
 * archivos externos en {@code salidas_referencia}.
 */
final class AiosTemplateService {

    private static final Logger log = LoggerFactory.getLogger(AiosTemplateService.class);
    private static final String CLASSPATH_DIR = "aios-templates/";

    private final AiosProperties properties;

    AiosTemplateService(AiosProperties properties) {
        this.properties = properties;
    }

    Workbook openWorkbook(String... candidateNames) throws IOException {
        if (candidateNames == null || candidateNames.length == 0) {
            throw new IllegalArgumentException("Debe indicar al menos un nombre de plantilla AIOS");
        }

        Path externalDir = properties == null ? null : properties.plantillaDir();
        Workbook external = tryOpenFromDirectory(externalDir, candidateNames, "plantilla externa");
        if (external != null) return external;

        ClassLoader classLoader = AiosTemplateService.class.getClassLoader();
        for (String name : candidateNames) {
            String resourceName = CLASSPATH_DIR + name;
            try (InputStream in = classLoader.getResourceAsStream(resourceName)) {
                if (in == null) continue;
                log.info("Plantilla AIOS interna seleccionada: classpath:{}", resourceName);
                return WorkbookFactory.create(in);
            }
        }

        // Compatibilidad con instalaciones anteriores. Esta ruta ya no es obligatoria.
        Path legacyDir = properties == null ? null : properties.salidasReferenciaDir();
        Workbook legacy = tryOpenFromDirectory(legacyDir, candidateNames, "referencia heredada");
        if (legacy != null) return legacy;

        throw new IllegalStateException("No se encontró plantilla AIOS. Se buscaron "
                + Arrays.toString(candidateNames)
                + " en aios.plantilla-dir, recursos internos y, opcionalmente, salidas_referencia");
    }

    private Workbook tryOpenFromDirectory(Path directory, String[] names, String source) throws IOException {
        if (directory == null) return null;
        for (String name : names) {
            Path candidate = directory.resolve(name);
            if (!Files.isRegularFile(candidate)) continue;
            log.info("Plantilla AIOS seleccionada desde {}: {}", source, candidate.toAbsolutePath());
            try (InputStream in = Files.newInputStream(candidate)) {
                return WorkbookFactory.create(in);
            }
        }
        return null;
    }
}
