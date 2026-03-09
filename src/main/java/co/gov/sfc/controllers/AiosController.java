package co.gov.sfc.controllers;

import co.gov.sfc.model.ModoGeneracion;
import co.gov.sfc.services.AiosGeneracionService;
import org.springframework.core.io.FileSystemResource;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.ExceptionHandler;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;

import static org.springframework.http.MediaType.TEXT_PLAIN;

import java.time.LocalDate;

@RestController
@RequestMapping("/aios")
public class AiosController {

    private final AiosGeneracionService generacionService;

    public AiosController(AiosGeneracionService generacionService) {
        this.generacionService = generacionService;
    }

    @PostMapping("/generar")
    public ResponseEntity<FileSystemResource> generar(
            @RequestParam LocalDate fechaCorte,
            @RequestParam ModoGeneracion modo
    ) {
        var resultado = generacionService.generar(fechaCorte, modo);
        var archivo = resultado.archivosGenerados().getFirst();
        var mediaType = resultado.zip() ? "application/zip" : "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=\"" + archivo.getFileName() + "\"")
                .contentType(MediaType.parseMediaType(mediaType))
                .body(new FileSystemResource(archivo));
    }
    @ExceptionHandler(Exception.class)
    public ResponseEntity<String> handleException(Exception ex) {
        return ResponseEntity.internalServerError()
                .contentType(TEXT_PLAIN)
                .body("Error al generar archivo AIOS: " + ex.getMessage());
    }

}
