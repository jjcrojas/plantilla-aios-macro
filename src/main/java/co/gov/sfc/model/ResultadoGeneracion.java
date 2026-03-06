package co.gov.sfc.model;

import java.nio.file.Path;
import java.util.List;

public record ResultadoGeneracion(List<Path> archivosGenerados, boolean zip) {
}
