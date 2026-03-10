package co.gov.sfc.controllers;

import co.gov.sfc.services.AiosGeneracionService;
import org.junit.jupiter.api.Test;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.test.autoconfigure.web.servlet.WebMvcTest;
import org.springframework.boot.test.mock.mockito.MockBean;
import org.springframework.test.web.servlet.MockMvc;

import static org.mockito.ArgumentMatchers.any;
import static org.mockito.Mockito.when;
import static org.springframework.test.web.servlet.request.MockMvcRequestBuilders.post;
import static org.springframework.test.web.servlet.result.MockMvcResultMatchers.content;
import static org.springframework.test.web.servlet.result.MockMvcResultMatchers.status;

@WebMvcTest(AiosController.class)
class AiosControllerErrorHandlingTest {

    @Autowired
    private MockMvc mockMvc;

    @MockBean
    private AiosGeneracionService generacionService;

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
