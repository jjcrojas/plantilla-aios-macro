package co.gov.sfc.util;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import org.junit.jupiter.api.Test;
import org.springframework.ui.ExtendedModelMap;

import co.gov.sfc.controllers.VersionModelAdvice;

class VersionAplicacionTest {

    @Test
    void vistaComparteVersionFechaYNombreDeCompilacion() {
        var model = new ExtendedModelMap();
        new VersionModelAdvice().agregarVersion(model);

        assertTrue(VersionAplicacion.version().matches("[0-9]+\\.[0-9]+"));
        assertTrue(VersionAplicacion.fecha().matches("\\d{2}/\\d{2}/\\d{4} \\d{2}:\\d{2}"));
        assertEquals("Generador AIOS - Delegatura para Pensiones", VersionAplicacion.nombre());
        assertEquals(VersionAplicacion.version(), model.get("versionAplicacion"));
        assertEquals(VersionAplicacion.fecha(), model.get("fechaVersion"));
        assertEquals(VersionAplicacion.nombre(), model.get("nombreAplicacion"));
        assertEquals("Versión " + VersionAplicacion.version()
                + " | Compilación " + VersionAplicacion.fecha()
                + " | App: " + VersionAplicacion.nombre(), VersionAplicacion.pie());
    }
}
