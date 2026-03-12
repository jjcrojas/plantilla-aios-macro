package co.gov.sfc.services;

import co.gov.sfc.excel.MensualDataReader;
import co.gov.sfc.excel.MensualExcelGenerator;
import co.gov.sfc.model.ModoGeneracion;
import co.gov.sfc.model.ResultadoGeneracion;
import org.springframework.stereotype.Service;

import java.io.IOException;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.LocalDate;
import java.util.ArrayList;
import java.util.List;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

@Service
public class AiosGeneracionService {

    private final MensualDataReader mensualDataReader;
    private final MensualExcelGenerator mensualExcelGenerator;

    public AiosGeneracionService(MensualDataReader mensualDataReader, MensualExcelGenerator mensualExcelGenerator) {
        this.mensualDataReader = mensualDataReader;
        this.mensualExcelGenerator = mensualExcelGenerator;
    }

    public ResultadoGeneracion generar(LocalDate fechaCorte, ModoGeneracion modo) {
        List<Path> archivos = new ArrayList<>();

        try {
            if (modo == ModoGeneracion.MENSUAL || modo == ModoGeneracion.TODO) {
                var mensual = mensualExcelGenerator.generar(mensualDataReader.read(fechaCorte));
                archivos.add(mensual);
            }
        } catch (OutOfMemoryError oom) {
            throw new IllegalStateException("Memoria insuficiente generando AIOS. Intente con más heap (-Xmx) o reduzca insumos cargados.", oom);
        }

        if (modo == ModoGeneracion.TRIMESTRAL || modo == ModoGeneracion.SEMESTRAL) {
            throw new UnsupportedOperationException("TRIMESTRAL/SEMESTRAL se dejarán para la siguiente iteración de migración");
        }

        if (modo == ModoGeneracion.TODO) {
            Path zip = zip(archivos);
            return new ResultadoGeneracion(List.of(zip), true);
        }
        return new ResultadoGeneracion(archivos, false);
    }

    private Path zip(List<Path> archivos) {
        Path zip = Path.of("target", "aios-output", "aios-generados.zip");
        try {
            Files.createDirectories(zip.getParent());
            try (OutputStream os = Files.newOutputStream(zip); ZipOutputStream zos = new ZipOutputStream(os)) {
                for (Path archivo : archivos) {
                    zos.putNextEntry(new ZipEntry(archivo.getFileName().toString()));
                    Files.copy(archivo, zos);
                    zos.closeEntry();
                }
            }
            return zip;
        } catch (IOException e) {
            throw new IllegalStateException("No fue posible crear ZIP", e);
        }
    }
}
