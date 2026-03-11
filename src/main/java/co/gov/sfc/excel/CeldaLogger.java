package co.gov.sfc.excel;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Component;

@Component
public class CeldaLogger {

    private static final Logger log = LoggerFactory.getLogger("AIOS_CELDAS");

    public void log(String hoja, String celda, Object valor, String archivoFuente, String hojaFuente, String celdaFuente) {
        log.info("hoja={}, celda={}, valor={}, fuente={}/{}/{}", hoja, celda, valor, archivoFuente, hojaFuente, celdaFuente);
    }
}
