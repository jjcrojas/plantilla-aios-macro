package co.gov.sfc.model;

public record CeldaTrazabilidad(
        String hojaDestino,
        String celdaDestino,
        Object valor,
        String fuenteArchivo,
        String fuenteHoja,
        String fuenteCelda
) {
}
