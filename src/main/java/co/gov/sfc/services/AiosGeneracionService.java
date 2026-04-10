package co.gov.sfc.services;

import co.gov.sfc.excel.MensualDataReader;
import co.gov.sfc.excel.MensualExcelGenerator;
import co.gov.sfc.excel.SemestralExcelGenerator;
import co.gov.sfc.excel.TrimestralDataReader;
import co.gov.sfc.excel.TrimestralExcelGenerator;
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
    private final SemestralExcelGenerator semestralExcelGenerator;
    private final TrimestralDataReader trimestralDataReader;
    private final TrimestralExcelGenerator trimestralExcelGenerator;

    public AiosGeneracionService(MensualDataReader mensualDataReader,
                                 MensualExcelGenerator mensualExcelGenerator,
                                 SemestralExcelGenerator semestralExcelGenerator,
                                 TrimestralDataReader trimestralDataReader,
                                 TrimestralExcelGenerator trimestralExcelGenerator) {
        this.mensualDataReader = mensualDataReader;
        this.mensualExcelGenerator = mensualExcelGenerator;
        this.semestralExcelGenerator = semestralExcelGenerator;
        this.trimestralDataReader = trimestralDataReader;
        this.trimestralExcelGenerator = trimestralExcelGenerator;
    }

    public ResultadoGeneracion generar(LocalDate fechaCorte, ModoGeneracion modo) {
        List<Path> archivos = new ArrayList<>();

        try {
            if (modo == ModoGeneracion.MENSUAL || modo == ModoGeneracion.TODO) {
                var mensual = mensualExcelGenerator.generar(mensualDataReader.read(fechaCorte));
                archivos.add(mensual);
            }

            if (modo == ModoGeneracion.TRIMESTRAL && !isQuarterMonth(fechaCorte)) {
                throw new IllegalArgumentException("La generación trimestral solo aplica para cortes de marzo, junio, septiembre o diciembre");
            }

            if (modo == ModoGeneracion.TRIMESTRAL || (modo == ModoGeneracion.TODO && isQuarterMonth(fechaCorte))) {
                var trimestral = trimestralExcelGenerator.generar(fechaCorte, trimestralDataReader.read(fechaCorte));
                archivos.add(trimestral);
            }

            if (modo == ModoGeneracion.SEMESTRAL && !isSemesterMonth(fechaCorte)) {
                throw new IllegalArgumentException("La generación semestral solo aplica para cortes de junio o diciembre");
            }

            if (modo == ModoGeneracion.SEMESTRAL || (modo == ModoGeneracion.TODO && isSemesterMonth(fechaCorte))) {
                var mensual = mensualDataReader.read(fechaCorte);
                var trimestral = trimestralDataReader.read(fechaCorte);
                var semestral = semestralExcelGenerator.generar(fechaCorte, mensual, trimestral);
                archivos.add(semestral);
            }
        } catch (OutOfMemoryError oom) {
            throw new IllegalStateException("Memoria insuficiente generando AIOS. Intente con más heap (-Xmx) o reduzca insumos cargados.", oom);
        }

        if (modo == ModoGeneracion.TODO) {
            Path zip = zip(archivos);
            return new ResultadoGeneracion(List.of(zip), true);
        }
        return new ResultadoGeneracion(archivos, false);
    }

    private boolean isQuarterMonth(LocalDate fechaCorte) {
        int m = fechaCorte.getMonthValue();
        return m == 3 || m == 6 || m == 9 || m == 12;
    }

    private boolean isSemesterMonth(LocalDate fechaCorte) {
        int m = fechaCorte.getMonthValue();
        return m == 6 || m == 12;
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
