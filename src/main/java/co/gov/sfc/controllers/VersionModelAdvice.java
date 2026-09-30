package co.gov.sfc.controllers;

import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.ControllerAdvice;
import org.springframework.web.bind.annotation.ModelAttribute;

import co.gov.sfc.util.VersionAplicacion;

@ControllerAdvice
public class VersionModelAdvice {

    @ModelAttribute
    public void agregarVersion(Model model) {
        model.addAttribute("versionAplicacion", VersionAplicacion.version());
        model.addAttribute("fechaVersion", VersionAplicacion.fecha());
        model.addAttribute("nombreAplicacion", VersionAplicacion.nombre());
    }
}
