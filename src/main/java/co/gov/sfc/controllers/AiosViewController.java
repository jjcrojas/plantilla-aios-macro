package co.gov.sfc.controllers;

import co.gov.sfc.model.ModoGeneracion;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.GetMapping;

import java.time.LocalDate;

@Controller
public class AiosViewController {

    @GetMapping({"/", "/ui", "/aios"})
    public String index(Model model) {
        model.addAttribute("modos", ModoGeneracion.values());
        model.addAttribute("fechaHoy", LocalDate.now());
        return "aios-index";
    }

    @GetMapping("/aios/generar")
    public String generarHelp(Model model) {
        model.addAttribute("modos", ModoGeneracion.values());
        model.addAttribute("fechaHoy", LocalDate.now());
        model.addAttribute("mensaje", "Use este formulario para enviar POST a /aios/generar.");
        return "aios-index";
    }
}
