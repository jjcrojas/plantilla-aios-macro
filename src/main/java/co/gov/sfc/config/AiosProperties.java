package co.gov.sfc.config;

import org.springframework.boot.context.properties.ConfigurationProperties;

import java.nio.file.Path;

@ConfigurationProperties(prefix = "aios")
public record AiosProperties(Path insumosDir, Path plantillaDir, Path salidasReferenciaDir, Integer maxPoiFileMb, Boolean macroRecalc491493) {
}
