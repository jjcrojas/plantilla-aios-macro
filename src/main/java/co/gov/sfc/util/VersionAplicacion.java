package co.gov.sfc.util;

import java.io.IOException;
import java.io.InputStream;
import java.io.UncheckedIOException;
import java.time.Instant;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.util.Properties;

/** Metadatos inmutables de la compilación, disponibles para todas las vistas. */
public final class VersionAplicacion {

    private static final String NOMBRE_PREDETERMINADO = "Generador AIOS - Delegatura para Pensiones";
    private static final Properties BUILD = cargar();

    private VersionAplicacion() {
    }

    private static Properties cargar() {
        Properties properties = new Properties();
        try (InputStream input = VersionAplicacion.class.getResourceAsStream("/META-INF/build-info.properties")) {
            if (input != null) {
                properties.load(input);
            }
        } catch (IOException error) {
            throw new UncheckedIOException("No se pudieron leer los metadatos de compilación", error);
        }
        return properties;
    }

    public static String version() {
        return BUILD.getProperty("build.version", "desarrollo local");
    }

    public static String fecha() {
        String compilation = BUILD.getProperty("build.compilation");
        if (compilation != null && !compilation.isBlank()) {
            return compilation;
        }
        String time = BUILD.getProperty("build.time");
        if (time == null || time.isBlank()) {
            return "no disponible";
        }
        return DateTimeFormatter.ofPattern("dd/MM/yyyy HH:mm")
                .withZone(ZoneId.of("America/Bogota"))
                .format(Instant.parse(time));
    }

    public static String nombre() {
        return BUILD.getProperty("build.applicationName", NOMBRE_PREDETERMINADO);
    }

    public static String pie() {
        return "Versión " + version() + " | Compilación " + fecha() + " | App: " + nombre();
    }
}
