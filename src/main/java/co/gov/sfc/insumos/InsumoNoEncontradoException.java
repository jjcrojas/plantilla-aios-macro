package co.gov.sfc.insumos;

/** Ausencia de un archivo de datos, distinta de un fallo de cálculo o configuración. */
public class InsumoNoEncontradoException extends IllegalArgumentException {
    public InsumoNoEncontradoException(String message) {
        super(message);
    }
}
