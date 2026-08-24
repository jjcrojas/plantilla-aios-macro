package co.gov.sfc.excel;

import org.apache.pdfbox.Loader;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.rendering.ImageType;
import org.apache.pdfbox.rendering.PDFRenderer;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Service;

import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.IOException;
import java.math.BigDecimal;
import java.net.URI;
import java.net.http.HttpClient;
import java.net.http.HttpRequest;
import java.net.http.HttpResponse;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.Duration;
import java.time.LocalDate;
import java.time.Month;
import java.text.Normalizer;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

@Service
public class ComisionesSfcService {

    private static final Logger log = LoggerFactory.getLogger(ComisionesSfcService.class);
    private static final URI BASE = URI.create("https://www.superfinanciera.gov.co");
    private static final URI INDEX = BASE.resolve("/publicaciones/20149");
    private static final Duration TIMEOUT = Duration.ofSeconds(90);
    private static final Pattern ROW = Pattern.compile("<tr\\b[^>]*>(.*?)</tr>", Pattern.CASE_INSENSITIVE | Pattern.DOTALL);
    private static final Pattern CELL = Pattern.compile("<td\\b[^>]*>(.*?)</td>", Pattern.CASE_INSENSITIVE | Pattern.DOTALL);
    private static final Pattern LINK = Pattern.compile("<a\\b[^>]*href=[\\\"']([^\\\"']+)[\\\"'][^>]*>(.*?)</a>", Pattern.CASE_INSENSITIVE | Pattern.DOTALL);
    private static final Pattern PERCENT = Pattern.compile("(\\d{1,2}[.,]\\d{1,2})\\s*%");
    private static final Pattern CORRUPTED_COMMISSION_BEFORE_INSURANCE =
            Pattern.compile("\\b(\\d{4,6})\\s+(\\d{1,2}[.,]\\d{2})\\s*%");
    private static final BigDecimal DISTRIBUCION_COMISION_SEGURO = new BigDecimal("3.00");

    private final HttpClient httpClient;
    private final Path tesseract;
    private final Path cacheRoot;

    public ComisionesSfcService() {
        this(HttpClient.newBuilder().connectTimeout(Duration.ofSeconds(15)).followRedirects(HttpClient.Redirect.NORMAL).build(),
                findTesseract(), Path.of("target", "aios-cache", "comisiones-sfc"));
    }

    ComisionesSfcService(HttpClient httpClient, Path tesseract, Path cacheRoot) {
        this.httpClient = httpClient;
        this.tesseract = tesseract;
        this.cacheRoot = cacheRoot;
    }

    public Map<String, BigDecimal> leer(LocalDate fechaCorte) {
        try {
            LocalDate publicacionEsperada = fechaCorte.plusMonths(1);
            String indexHtml = getText(INDEX);
            URI paginaAnual = findAnnualLettersPage(indexHtml, publicacionEsperada.getYear());
            String annualHtml = getText(paginaAnual);
            URI pdfUri = findCommissionPdf(annualHtml, publicacionEsperada.getMonth());
            byte[] pdf = getBytes(pdfUri);
            if (pdf.length == 0 || pdf.length > 20 * 1024 * 1024) {
                throw new IllegalStateException("El PDF de comisiones tiene un tamaño inválido: " + pdf.length);
            }

            Path workDir = cacheRoot.resolve(fechaCorte.toString());
            Files.createDirectories(workDir);
            Path pdfPath = workDir.resolve("carta-circular-comisiones.pdf");
            Files.write(pdfPath, pdf);
            String ocr = ocr(pdf, workDir);
            Files.writeString(workDir.resolve("ocr.txt"), ocr, StandardCharsets.UTF_8);
            validateCutoff(ocr, fechaCorte);
            Map<String, BigDecimal> values = parseCommissionTable(ocr);
            log.info("Carta Circular SFC procesada fechaCorte={} paginaAnual={} pdf={} valores={}",
                    fechaCorte, paginaAnual, pdfUri, values);
            return values;
        } catch (Exception e) {
            throw new IllegalStateException("No se pudieron obtener las comisiones de la Superfinanciera", e);
        }
    }

    private String getText(URI uri) throws IOException, InterruptedException {
        return new String(send(uri), StandardCharsets.UTF_8);
    }

    private byte[] getBytes(URI uri) throws IOException, InterruptedException {
        return send(uri);
    }

    private byte[] send(URI uri) throws IOException, InterruptedException {
        if (!"www.superfinanciera.gov.co".equalsIgnoreCase(uri.getHost())) {
            throw new IllegalArgumentException("Host SFC inesperado: " + uri);
        }
        HttpRequest request = HttpRequest.newBuilder(uri).timeout(TIMEOUT)
                .header("User-Agent", "plantilla-aios-generator/1.0")
                .GET().build();
        HttpResponse<byte[]> response = httpClient.send(request, HttpResponse.BodyHandlers.ofByteArray());
        if (response.statusCode() < 200 || response.statusCode() >= 300) {
            throw new IllegalStateException("HTTP " + response.statusCode() + " consultando " + uri);
        }
        return response.body();
    }

    static URI findAnnualLettersPage(String html, int year) {
        int header = normalize(html).indexOf("cartas circulares (2)");
        if (header < 0) throw new IllegalStateException("No se encontró la columna Cartas Circulares en el índice SFC");
        int tableEnd = html.indexOf("</table>", header);
        String table = html.substring(header, tableEnd < 0 ? html.length() : tableEnd);
        Matcher rows = ROW.matcher(table);
        while (rows.find()) {
            List<String> cells = cells(rows.group(1));
            if (cells.size() < 3 || !stripTags(cells.get(1)).equals(String.valueOf(year))) continue;
            Matcher link = LINK.matcher(cells.get(1));
            if (link.find()) return resolveSfc(link.group(1));
        }
        throw new IllegalStateException("No se encontró la página de Cartas Circulares para " + year);
    }

    static URI findCommissionPdf(String html, Month expectedPublicationMonth) {
        Matcher rows = ROW.matcher(html);
        while (rows.find()) {
            List<String> cells = cells(rows.group(1));
            if (cells.size() < 3) continue;
            String date = normalize(stripTags(cells.get(1)));
            String subject = normalize(stripTags(cells.get(2)));
            if (!date.startsWith(spanishMonth(expectedPublicationMonth))) continue;
            if (!(subject.contains("rentabilidad") && subject.contains("comision de administracion") && subject.contains("seguro previsional"))) continue;
            Matcher link = LINK.matcher(cells.get(0));
            if (link.find()) return resolveSfc(link.group(1));
        }
        throw new IllegalStateException("No se encontró la Carta Circular de comisiones publicada en " + expectedPublicationMonth);
    }

    private String ocr(byte[] pdf, Path workDir) throws Exception {
        if (!Files.isRegularFile(tesseract)) {
            throw new IllegalStateException("No se encontró Tesseract OCR. Configure TESSERACT_PATH o instálelo en AppData\\Local\\Tesseract-OCR");
        }
        StringBuilder all = new StringBuilder();
        try (PDDocument document = Loader.loadPDF(pdf)) {
            PDFRenderer renderer = new PDFRenderer(document);
            for (int page = 0; page < document.getNumberOfPages(); page++) {
                Path image = workDir.resolve("pagina-" + (page + 1) + ".png");
                BufferedImage rendered = renderer.renderImageWithDPI(page, 300, ImageType.GRAY);
                ImageIO.write(rendered, "png", image.toFile());
                if (page == 0) {
                    int x = (int) (rendered.getWidth() * 0.10);
                    int y = (int) (rendered.getHeight() * 0.40);
                    int width = (int) (rendered.getWidth() * 0.55);
                    int height = (int) (rendered.getHeight() * 0.21);
                    Path table = workDir.resolve("tabla-comision-seguro.png");
                    ImageIO.write(rendered.getSubimage(x, y, width, height), "png", table.toFile());
                    all.append(runTesseract(table)).append('\n');
                }
                all.append(runTesseract(image)).append('\n');
            }
        }
        return all.toString();
    }

    private String runTesseract(Path image) throws IOException, InterruptedException {
        Process process = new ProcessBuilder(tesseract.toString(), image.toString(), "stdout", "-l", "eng", "--psm", "6")
                .redirectErrorStream(true).start();
        String text = new String(process.getInputStream().readAllBytes(), StandardCharsets.UTF_8);
        if (process.waitFor() != 0) throw new IllegalStateException("Tesseract terminó con error: " + text);
        return text;
    }

    static Map<String, BigDecimal> parseCommissionTable(String ocr) {
        Map<String, BigDecimal> out = new HashMap<>();
        Matcher otherCommissions = Pattern.compile("OTRAS\\s+COMISIONES\\s+AUTORIZADAS", Pattern.CASE_INSENSITIVE).matcher(ocr);
        String contributionDistribution = otherCommissions.find() ? ocr.substring(0, otherCommissions.start()) : ocr;
        for (String line : contributionDistribution.split("\\R")) {
            String normalized = normalize(line);
            String prefix = adminPrefix(normalized);
            if (prefix == null || out.containsKey(prefix + "_obl")) continue;
            Matcher values = PERCENT.matcher(line);
            List<BigDecimal> percentages = new ArrayList<>();
            while (values.find() && percentages.size() < 2) {
                percentages.add(new BigDecimal(values.group(1).replace(',', '.')));
            }
            BigDecimal commission;
            BigDecimal insurance;
            Matcher corrupted = CORRUPTED_COMMISSION_BEFORE_INSURANCE.matcher(line);
            if (corrupted.find()) {
                insurance = new BigDecimal(corrupted.group(2).replace(',', '.'));
                commission = DISTRIBUCION_COMISION_SEGURO.subtract(insurance);
                String expectedDigits = commission.setScale(2).toPlainString().replace(".", "");
                if (!corrupted.group(1).startsWith(expectedDigits)) continue;
                log.warn("OCR reconstruyó comisión de {} desde token corrupto {} y seguro {}: {}",
                        prefix, corrupted.group(1), insurance, commission);
            } else if (percentages.size() >= 2) {
                commission = percentages.get(0);
                insurance = percentages.get(1);
            } else {
                continue;
            }
            BigDecimal sum = commission.add(insurance);
            if (sum.subtract(DISTRIBUCION_COMISION_SEGURO).abs().compareTo(new BigDecimal("0.02")) > 0) {
                if (!validPercentage(commission) && validPercentage(insurance)) {
                    BigDecimal corrected = DISTRIBUCION_COMISION_SEGURO.subtract(insurance);
                    log.warn("OCR corrigió comisión de {}: leído={} seguro={} corregido={}", prefix, commission, insurance, corrected);
                    commission = corrected;
                } else if (validPercentage(commission) && !validPercentage(insurance)) {
                    insurance = DISTRIBUCION_COMISION_SEGURO.subtract(commission);
                } else {
                    continue;
                }
            }
            if (!validPercentage(commission) || !validPercentage(insurance)) continue;
            out.put(prefix + "_obl", commission);
            out.put(prefix + "_seg", insurance);
        }
        List<String> required = List.of("col_obl", "col_seg", "por_obl", "por_seg", "pro_obl", "pro_seg", "ska_obl", "ska_seg");
        List<String> missing = required.stream().filter(k -> !out.containsKey(k)).toList();
        if (!missing.isEmpty()) throw new IllegalStateException("OCR incompleto; faltan valores " + missing);
        return Map.copyOf(out);
    }

    private static void validateCutoff(String ocr, LocalDate cutoff) {
        String expected = cutoff.getDayOfMonth() + " de " + spanishMonth(cutoff.getMonth()) + " de " + cutoff.getYear();
        if (!normalize(ocr).contains(expected)) {
            throw new IllegalStateException("La carta descargada no declara el corte esperado " + expected);
        }
    }

    private static boolean validPercentage(BigDecimal value) {
        return value.signum() >= 0 && value.compareTo(DISTRIBUCION_COMISION_SEGURO) <= 0;
    }

    private static String adminPrefix(String line) {
        if (line.contains("colfondos") || line.contains("colfonbos")) return "col";
        if (line.contains("porvenir")) return "por";
        if (line.contains("proteccion")) return "pro";
        if (line.contains("skandia")) return "ska";
        return null;
    }

    private static List<String> cells(String row) {
        List<String> cells = new ArrayList<>();
        Matcher matcher = CELL.matcher(row);
        while (matcher.find()) cells.add(matcher.group(1));
        return cells;
    }

    private static String stripTags(String html) {
        return decodeHtmlEntities(html.replaceAll("(?is)<[^>]+>", " "))
                .replaceAll("\\s+", " ").trim();
    }

    private static String decodeHtmlEntities(String html) {
        return html.replace("&nbsp;", " ").replace("&#160;", " ")
                .replace("&aacute;", "á").replace("&eacute;", "é")
                .replace("&iacute;", "í").replace("&oacute;", "ó")
                .replace("&uacute;", "ú").replace("&ntilde;", "ñ")
                .replace("&Aacute;", "Á").replace("&Eacute;", "É")
                .replace("&Iacute;", "Í").replace("&Oacute;", "Ó")
                .replace("&Uacute;", "Ú").replace("&Ntilde;", "Ñ")
                .replace("&quot;", "\"").replace("&#39;", "'")
                .replace("&amp;", "&");
    }

    private static URI resolveSfc(String href) {
        return BASE.resolve(href.replace("&amp;", "&"));
    }

    private static String normalize(String value) {
        if (value == null) return "";
        return Normalizer.normalize(value, Normalizer.Form.NFD).replaceAll("\\p{M}", "")
                .toLowerCase(Locale.ROOT).replaceAll("\\s+", " ").trim();
    }

    private static String spanishMonth(Month month) {
        return switch (month) {
            case JANUARY -> "enero"; case FEBRUARY -> "febrero"; case MARCH -> "marzo";
            case APRIL -> "abril"; case MAY -> "mayo"; case JUNE -> "junio";
            case JULY -> "julio"; case AUGUST -> "agosto"; case SEPTEMBER -> "septiembre";
            case OCTOBER -> "octubre"; case NOVEMBER -> "noviembre"; case DECEMBER -> "diciembre";
        };
    }

    private static Path findTesseract() {
        String configured = System.getenv("TESSERACT_PATH");
        List<Path> candidates = new ArrayList<>();
        if (configured != null && !configured.isBlank()) candidates.add(Path.of(configured));
        candidates.add(Path.of(System.getProperty("user.home"), "AppData", "Local", "Tesseract-OCR", "tesseract.exe"));
        String programFiles = System.getenv("ProgramFiles");
        if (programFiles != null) candidates.add(Path.of(programFiles, "Tesseract-OCR", "tesseract.exe"));
        return candidates.stream().filter(Files::isRegularFile).findFirst().orElse(candidates.get(0));
    }
}
