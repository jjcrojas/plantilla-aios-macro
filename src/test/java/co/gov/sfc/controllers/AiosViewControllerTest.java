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
    void shouldRenderPeriodSelectorAndAllModes() throws Exception {
        mockMvc.perform(get("/aios"))
                .andExpect(status().isOk())
                .andExpect(org.springframework.test.web.servlet.result.MockMvcResultMatchers.content().string(
                        org.hamcrest.Matchers.allOf(
                                org.hamcrest.Matchers.containsString("id=\"tipoPeriodo\""),
                                org.hamcrest.Matchers.containsString("id=\"periodoInicial\""),
                                org.hamcrest.Matchers.containsString("id=\"periodoFinal\""),
                                org.hamcrest.Matchers.containsString("class=\"sfc-brand-header\""),
                                org.hamcrest.Matchers.containsString("class=\"card generator-card\""),
                                org.hamcrest.Matchers.containsString("Versión"),
                                org.hamcrest.Matchers.containsString("Generador AIOS - Delegatura para Pensiones"),
                                org.hamcrest.Matchers.containsString("value=\"MENSUAL\""),
                                org.hamcrest.Matchers.containsString("value=\"TRIMESTRAL\""),
                                org.hamcrest.Matchers.containsString("value=\"SEMESTRAL\""),
                                org.hamcrest.Matchers.containsString("value=\"TODO\""))));
    }

    @Test
    void shouldRenderUiAtAiosGenerarGet() throws Exception {
        mockMvc.perform(get("/aios/generar"))
                .andExpect(status().isOk())
                .andExpect(view().name("aios-index"));
    }
}
