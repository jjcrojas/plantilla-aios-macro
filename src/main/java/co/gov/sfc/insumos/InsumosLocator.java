package co.gov.sfc.insumos;

import co.gov.sfc.config.AiosProperties;
import org.springframework.stereotype.Component;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.LocalDate;
import java.time.format.TextStyle;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.List;
import java.util.Locale;

@Component
public class InsumosLocator {

    private static final Locale ES_CO = new Locale("es", "CO");
    private final AiosProperties properties;

    public InsumosLocator(AiosProperties properties) {
        this.properties = properties;
    }

    public Path findRequired(String contains) {
        return findRequired(contains, null);
    }

    public Path findRequired(String contains, LocalDate fechaCorte) {
        if (properties.insumosDir() == null) {
            throw new IllegalStateException("La propiedad 'aios.insumos-dir' no está configurada.");
        }

        List<Path> candidates = new ArrayList<>();
        if (fechaCorte != null) {
            candidates.addAll(candidateDirsFor(contains, fechaCorte));
        }
        candidates.add(properties.insumosDir());

        for (Path dir : candidates) {
            Path found = tryFind(dir, contains);
            if (found != null) {
                return found;
            }
        }

        throw new IllegalArgumentException("No se encontró insumo que contenga: " + contains + " en rutas: " + candidates);
    }

    private List<Path> candidateDirsFor(String contains, LocalDate fechaCorte) {
        Path base = properties.insumosDir();
        int year = fechaCorte.getYear();
        int month = fechaCorte.getMonthValue();
        String monthName = fechaCorte.getMonth().getDisplayName(TextStyle.FULL, ES_CO).toLowerCase();
        String monthFolderLower = month + " " + monthName;
        String monthFolderUpper = month + " " + capitalize(monthName);

        List<Path> dirs = new ArrayList<>();

        String normalized = contains.toLowerCase();
        if (normalized.contains("491")) {
            dirs.add(base.resolve("Formato 491").resolve("491 FORMATO TRANSMITIDO"));
        } else if (normalized.contains("sistema total")) {
            dirs.add(base.resolve("Balances").resolve(String.valueOf(year)).resolve(monthFolderLower));
            dirs.add(base.resolve("Balances").resolve(String.valueOf(year)).resolve(monthFolderUpper));
        } else if (normalized.contains("limites")) {
            dirs.add(base.resolve("LIMITES").resolve(String.valueOf(year)).resolve(monthFolderLower));
            dirs.add(base.resolve("LIMITES").resolve(String.valueOf(year)).resolve(monthFolderUpper));
        } else if (normalized.contains("rent_vr_uni_moderado")) {
            dirs.add(base.resolve("PROCESOS MENSUALES")
                    .resolve("Rentabilidad Minima")
                    .resolve("Historico_Rent_minima")
                    .resolve(String.valueOf(year))
                    .resolve(monthFolderLower));
            dirs.add(base.resolve("PROCESOS MENSUALES")
                    .resolve("Rentabilidad Minima")
                    .resolve("Historico_Rent_minima")
                    .resolve(String.valueOf(year))
                    .resolve(monthFolderUpper));
        } else if (normalized.contains("formato_136_meses")) {
            dirs.add(base.resolve("FORMATOS ACTUALIZADOS"));
            dirs.add(base);
        } else if (normalized.contains("plantilla aios") || normalized.contains("plantilla_aios")) {
            dirs.add(base.resolve("plantillas"));
            dirs.add(base.resolve("FORMATOS ACTUALIZADOS"));
            dirs.add(base);
        } else if (normalized.contains("pib_pea_trm_dg")) {
            dirs.add(base);
        } else if (normalized.contains("493")) {
            dirs.add(base);
        }

        return dirs;
    }

    private String capitalize(String text) {
        if (text == null || text.isBlank()) return text;
        return Character.toUpperCase(text.charAt(0)) + text.substring(1);
    }

    private Path tryFind(Path dir, String contains) {
        if (dir == null || !Files.isDirectory(dir)) {
            return null;
        }
        try (var stream = Files.list(dir)) {
            return stream
                    .filter(Files::isRegularFile)
                    .filter(p -> p.getFileName().toString().toLowerCase().contains(contains.toLowerCase()))
                    .sorted(Comparator.comparing(Path::toString))
                    .findFirst()
                    .orElse(null);
        } catch (IOException e) {
            return null;
        }
    }
}
