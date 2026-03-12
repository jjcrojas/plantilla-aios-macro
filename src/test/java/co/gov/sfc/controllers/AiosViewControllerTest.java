package co.gov.sfc.controllers;

import co.gov.sfc.AIOSApplication;
import org.junit.jupiter.api.Test;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.test.autoconfigure.web.servlet.WebMvcTest;
import org.springframework.test.context.ContextConfiguration;
import org.springframework.test.web.servlet.MockMvc;

import static org.springframework.test.web.servlet.request.MockMvcRequestBuilders.get;
import static org.springframework.test.web.servlet.result.MockMvcResultMatchers.status;
import static org.springframework.test.web.servlet.result.MockMvcResultMatchers.view;

@WebMvcTest(AiosViewController.class)
@ContextConfiguration(classes = AIOSApplication.class)
class AiosViewControllerTest {

    @Autowired
    private MockMvc mockMvc;

    @Test
    void shouldRenderUiAtAiosPath() throws Exception {
        mockMvc.perform(get("/aios"))
                .andExpect(status().isOk())
                .andExpect(view().name("aios-index"));
    }

    @Test
    void shouldRenderUiAtAiosGenerarGet() throws Exception {
        mockMvc.perform(get("/aios/generar"))
                .andExpect(status().isOk())
                .andExpect(view().name("aios-index"));
    }
}
