package co.gov.sfc.services;

import co.gov.sfc.excel.MensualDataReader;
import co.gov.sfc.excel.MensualExcelGenerator;
import co.gov.sfc.excel.SemestralExcelGenerator;
import co.gov.sfc.excel.TrimestralDataReader;
import co.gov.sfc.excel.TrimestralExcelGenerator;
import co.gov.sfc.model.ModoGeneracion;
import co.gov.sfc.model.ResultadoGeneracion;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import java.io.IOException;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.LocalDate;
import java.time.YearMonth;
import java.util.ArrayList;
import java.util.List;
import java.util.function.IntPredicate;
import java.util.function.Supplier;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

@Service
public class AiosGeneracionService {

    private static final Logger log = LoggerFactory.getLogger(AiosGeneracionService.class);

    private final MensualDataReader mensualDataReader;
    private final MensualExcelGenerator mensualExcelGenerator;
    private final SemestralExcelGenerator semestralExcelGenerator;
    private final TrimestralDataReader trimestralDataReader;
    private final TrimestralExcelGenerator trimestralExcelGenerator;
    private final Path outputDir;

    @Autowired
    public AiosGeneracionService(MensualDataReader mensualDataReader,
                                 MensualExcelGenerator mensualExcelGenerator,
                                 SemestralExcelGenerator semestralExcelGenerator,
                                 TrimestralDataReader trimestralDataReader,
                                 TrimestralExcelGenerator trimestralExcelGenerator) {
        this(mensualDataReader, mensualExcelGenerator, semestralExcelGenerator,
                trimestralDataReader, trimestralExcelGenerator, Path.of("target", "aios-output"));
    }

    AiosGeneracionService(MensualDataReader mensualDataReader,
                          MensualExcelGenerator mensualExcelGenerator,
                          SemestralExcelGenerator semestralExcelGenerator,
                          TrimestralDataReader trimestralDataReader,
                          TrimestralExcelGenerator trimestralExcelGenerator,
                          Path outputDir) {
        this.mensualDataReader = mensualDataReader;
        this.mensualExcelGenerator = mensualExcelGenerator;
        this.semestralExcelGenerator = semestralExcelGenerator;
        this.trimestralDataReader = trimestralDataReader;
        this.trimestralExcelGenerator = trimestralExcelGenerator;
        this.outputDir = outputDir;
    }

    public ResultadoGeneracion generar(LocalDate fechaCorte, ModoGeneracion modo) {
        long start = System.currentTimeMillis();
        log.info("Inicio generación AIOS: fechaCorte={}, modo={}", fechaCorte, modo);
        List<Path> archivos = new ArrayList<>();

        try {
            if (modo == ModoGeneracion.TODO) {
                return generarTodoTolerante(fechaCorte, start);
            }

            if (modo == ModoGeneracion.MENSUAL) {
                var mensual = mensualExcelGenerator.generar(mensualDataReader.read(fechaCorte));
                archivos.add(mensual);
            }

            if (modo == ModoGeneracion.TRIMESTRAL && !isQuarterMonth(fechaCorte)) {
                throw new IllegalArgumentException("La generación trimestral solo aplica para cortes de marzo, junio, septiembre o diciembre");
            }

            if (modo == ModoGeneracion.TRIMESTRAL) {
                var trimestral = trimestralExcelGenerator.generar(fechaCorte, trimestralDataReader.read(fechaCorte));
                archivos.add(trimestral);
            }

            if (modo == ModoGeneracion.SEMESTRAL && !isSemesterMonth(fechaCorte)) {
                throw new IllegalArgumentException("La generación semestral solo aplica para cortes de junio o diciembre");
            }

            if (modo == ModoGeneracion.SEMESTRAL) {
                var mensual = mensualDataReader.read(fechaCorte);
                var trimestral = trimestralDataReader.readForSemestral(fechaCorte, mensual);
                var semestral = semestralExcelGenerator.generar(fechaCorte, mensual, trimestral);
                archivos.add(semestral);
            }
        } catch (OutOfMemoryError oom) {
            throw new IllegalStateException("Memoria insuficiente generando AIOS. Intente con más heap (-Xmx) o reduzca insumos cargados.", oom);
        }

        log.info("Generación AIOS finalizada en {} ms. Archivos generados={}", (System.currentTimeMillis() - start), archivos);
        return new ResultadoGeneracion(archivos, false);
    }

    private ResultadoGeneracion generarTodoTolerante(LocalDate fechaCorte, long start) {
        List<Path> archivos = new ArrayList<>();
        List<String> errores = new ArrayList<>();

        intentarGenerar("mensual", fechaCorte, archivos, errores,
                () -> mensualExcelGenerator.generar(mensualDataReader.read(fechaCorte)));

        if (isQuarterMonth(fechaCorte)) {
            intentarGenerar("trimestral", fechaCorte, archivos, errores,
                    () -> trimestralExcelGenerator.generar(fechaCorte, trimestralDataReader.read(fechaCorte)));
        }

        if (isSemesterMonth(fechaCorte)) {
            intentarGenerar("semestral", fechaCorte, archivos, errores, () -> {
                var mensual = mensualDataReader.read(fechaCorte);
                var trimestral = trimestralDataReader.readForSemestral(fechaCorte, mensual);
                return semestralExcelGenerator.generar(fechaCorte, mensual, trimestral);
            });
        }

        if (archivos.isEmpty()) {
            throw new IllegalStateException("No fue posible generar ningún archivo AIOS. "
                    + String.join(" | ", errores));
        }

        Path zip = zip(archivos);
        if (errores.isEmpty()) {
            log.info("Generación AIOS completa finalizada en {} ms. Archivo ZIP={}",
                    (System.currentTimeMillis() - start), zip.toAbsolutePath());
        } else {
            log.warn("Generación AIOS parcial finalizada en {} ms. Archivo ZIP={} archivosGenerados={} archivosOmitidos={}",
                    (System.currentTimeMillis() - start), zip.toAbsolutePath(), archivos, errores);
        }
        return new ResultadoGeneracion(List.of(zip), true);
    }

    private void intentarGenerar(String tipo,
                                 LocalDate fechaCorte,
                                 List<Path> archivos,
                                 List<String> errores,
                                 Supplier<Path> generador) {
        try {
            Path archivo = generador.get();
            archivos.add(archivo);
            log.info("Archivo {} generado correctamente para fechaCorte={}: {}",
                    tipo, fechaCorte, archivo.toAbsolutePath());
        } catch (Exception ex) {
            String detalle = tipo + ": " + ex.getMessage();
            errores.add(detalle);
            log.warn("No fue posible generar el archivo {} para fechaCorte={}; se continuará con los demás archivos. Causa: {}",
                    tipo, fechaCorte, ex.getMessage(), ex);
        }
    }

    public ResultadoGeneracion generarMensuales(LocalDate desde, LocalDate hasta) {
        validateRange(desde, hasta);

        List<co.gov.sfc.excel.MensualData> periodos = new ArrayList<>();
        YearMonth actual = YearMonth.from(desde);
        YearMonth ultimo = YearMonth.from(hasta);
        while (!actual.isAfter(ultimo)) {
            periodos.add(mensualDataReader.read(actual.atEndOfMonth()));
            actual = actual.plusMonths(1);
        }

        Path archivo = mensualExcelGenerator.generar(periodos);
        log.info("Generación mensual acumulada finalizada: desde={}, hasta={}, salida={}", desde, hasta, archivo.toAbsolutePath());
        return new ResultadoGeneracion(List.of(archivo), false);
    }

    public ResultadoGeneracion generarRango(LocalDate desde, LocalDate hasta, ModoGeneracion modo) {
        validateRange(desde, hasta);
        return switch (modo) {
            case MENSUAL -> generarMensuales(desde, hasta);
            case TRIMESTRAL -> generarTrimestrales(desde, hasta);
            case SEMESTRAL -> generarSemestrales(desde, hasta);
            case TODO -> throw new IllegalArgumentException(
                    "La generación consolidada por rango requiere modo MENSUAL, TRIMESTRAL o SEMESTRAL");
        };
    }

    public ResultadoGeneracion generarTrimestrales(LocalDate desde, LocalDate hasta) {
        validateRange(desde, hasta);
        List<LocalDate> fechas = cutoffs(desde, hasta, this::isQuarterMonthValue, "trimestral");
        List<TrimestralExcelGenerator.PeriodoTrimestral> periodos = new ArrayList<>();
        for (LocalDate fecha : fechas) {
            periodos.add(new TrimestralExcelGenerator.PeriodoTrimestral(
                    fecha, trimestralDataReader.read(fecha)));
        }
        Path archivo = trimestralExcelGenerator.generar(periodos);
        log.info("Generación trimestral consolidada finalizada: desde={}, hasta={}, periodos={}, salida={}",
                desde, hasta, fechas, archivo.toAbsolutePath());
        return new ResultadoGeneracion(List.of(archivo), false);
    }

    public ResultadoGeneracion generarSemestrales(LocalDate desde, LocalDate hasta) {
        validateRange(desde, hasta);
        List<LocalDate> fechas = cutoffs(desde, hasta, this::isSemesterMonthValue, "semestral");
        List<SemestralExcelGenerator.PeriodoSemestral> periodos = new ArrayList<>();
        for (LocalDate fecha : fechas) {
            var mensual = mensualDataReader.read(fecha);
            var trimestral = trimestralDataReader.readForSemestral(fecha, mensual);
            periodos.add(new SemestralExcelGenerator.PeriodoSemestral(fecha, mensual, trimestral));
        }
        Path archivo = semestralExcelGenerator.generar(periodos);
        log.info("Generación semestral consolidada finalizada: desde={}, hasta={}, periodos={}, salida={}",
                desde, hasta, fechas, archivo.toAbsolutePath());
        return new ResultadoGeneracion(List.of(archivo), false);
    }

    private void validateRange(LocalDate desde, LocalDate hasta) {
        if (desde.isAfter(hasta)) {
            throw new IllegalArgumentException("La fecha inicial no puede ser posterior a la fecha final");
        }
    }

    private List<LocalDate> cutoffs(LocalDate desde, LocalDate hasta, IntPredicate allowedMonth, String tipo) {
        List<LocalDate> fechas = new ArrayList<>();
        YearMonth actual = YearMonth.from(desde);
        YearMonth ultimo = YearMonth.from(hasta);
        while (!actual.isAfter(ultimo)) {
            LocalDate fecha = actual.atEndOfMonth();
            if (allowedMonth.test(actual.getMonthValue())
                    && !fecha.isBefore(desde)
                    && !fecha.isAfter(hasta)) {
                fechas.add(fecha);
            }
            actual = actual.plusMonths(1);
        }
        if (fechas.isEmpty()) {
            throw new IllegalArgumentException("El rango no contiene ningún corte " + tipo + " válido");
        }
        return fechas;
    }

    private boolean isQuarterMonth(LocalDate fechaCorte) {
        return isQuarterMonthValue(fechaCorte.getMonthValue());
    }

    private boolean isSemesterMonth(LocalDate fechaCorte) {
        return isSemesterMonthValue(fechaCorte.getMonthValue());
    }

    private boolean isQuarterMonthValue(int month) {
        return month == 3 || month == 6 || month == 9 || month == 12;
    }

    private boolean isSemesterMonthValue(int month) {
        return month == 6 || month == 12;
    }

    private Path zip(List<Path> archivos) {
        Path zip = outputDir.resolve("aios-generados.zip");
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
