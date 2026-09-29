package co.gov.sfc.controllers;

import static org.mockito.ArgumentMatchers.any;
import static org.mockito.Mockito.when;
import static org.springframework.test.web.servlet.request.MockMvcRequestBuilders.post;
import static org.springframework.test.web.servlet.result.MockMvcResultMatchers.content;
import static org.springframework.test.web.servlet.result.MockMvcResultMatchers.status;

import org.junit.jupiter.api.Test;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.test.autoconfigure.web.servlet.WebMvcTest;
import org.springframework.test.context.ContextConfiguration;
import org.springframework.test.context.bean.override.mockito.MockitoBean;
import org.springframework.test.web.servlet.MockMvc;

import co.gov.sfc.AIOSApplication;
import co.gov.sfc.services.AiosGeneracionService;

@WebMvcTest(AiosController.class)
@ContextConfiguration(classes = AIOSApplication.class)
class AiosControllerErrorHandlingTest {

    @Autowired
    private MockMvc mockMvc;

    @MockitoBean
    private AiosGeneracionService generacionService;

    @org.junit.jupiter.api.io.TempDir
    java.nio.file.Path tempDir;

    @org.junit.jupiter.params.ParameterizedTest
    @org.junit.jupiter.params.provider.EnumSource(co.gov.sfc.model.ModoGeneracion.class)
    void shouldDownloadRangeForEveryMode(co.gov.sfc.model.ModoGeneracion modo) throws Exception {
        boolean zip = modo == co.gov.sfc.model.ModoGeneracion.TODO;
        var file = java.nio.file.Files.writeString(tempDir.resolve(zip ? "aios.zip" : "aios.xlsx"), "test");
        var desde = java.time.LocalDate.of(2025, 6, 1);
        var hasta = java.time.LocalDate.of(2025, 12, 31);
        when(generacionService.generarRango(desde, hasta, modo))
                .thenReturn(new co.gov.sfc.model.ResultadoGeneracion(java.util.List.of(file), zip));
        mockMvc.perform(post("/aios/generar-rango")
                        .param("desde", desde.toString()).param("hasta", hasta.toString())
                        .param("modo", modo.name()))
                .andExpect(status().isOk())
                .andExpect(content().contentType(zip ? "application/zip"
                        : "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"));
        org.mockito.Mockito.verify(generacionService).generarRango(desde, hasta, modo);
    }

    @Test
    void shouldReturnPlainMessageWhenGenerationFails() throws Exception {
        when(generacionService.generar(any(), any())).thenThrow(new IllegalStateException("fallo controlado"));

        mockMvc.perform(post("/aios/generar")
                        .param("fechaCorte", "2025-06-30")
                        .param("modo", "MENSUAL"))
                .andExpect(status().isInternalServerError())
                .andExpect(content().string(org.hamcrest.Matchers.containsString("Error al generar archivo AIOS")));
    }
}
