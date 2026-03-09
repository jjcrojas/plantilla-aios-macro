package co.gov.sfc.controllers;

import co.gov.sfc.model.ModoGeneracion;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.GetMapping;

import java.time.LocalDate;

@Controller
public class AiosViewController {

    @GetMapping({"/", "/ui"})
    public String index(Model model) {
        model.addAttribute("modos", ModoGeneracion.values());
        model.addAttribute("fechaHoy", LocalDate.now());
        return "aios-index";
    }
}
