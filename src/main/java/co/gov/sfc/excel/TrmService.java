package co.gov.sfc.excel;

import co.gov.sfc.config.AiosProperties;
import co.gov.sfc.insumos.InsumosLocator;
import org.apache.poi.ss.usermodel.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import javax.xml.parsers.DocumentBuilderFactory;
import java.io.ByteArrayInputStream;
import java.math.BigDecimal;
import java.net.URI;
import java.net.http.HttpClient;
import java.net.http.HttpRequest;
import java.net.http.HttpResponse;
import java.nio.charset.StandardCharsets;
import java.time.Duration;
import java.time.LocalDate;
import java.util.Map;
import java.util.concurrent.ConcurrentHashMap;

/** Obtiene la TRM una sola vez por fecha y conserva el archivo histórico como contingencia. */
@Service
public class TrmService {

    private static final Logger log = LoggerFactory.getLogger(TrmService.class);
    private static final String SOAP_NAMESPACE =
            "http://action.trm.services.generic.action.superfinanciera.nexura.sc.com.co/";

    private final InsumosLocator locator;
    private final AiosProperties properties;
    private final HttpClient httpClient;
    private final URI endpoint;
    private final Duration requestTimeout;
    private final Map<LocalDate, BigDecimal> cache = new ConcurrentHashMap<>();

    @Autowired
    public TrmService(InsumosLocator locator,
                      AiosProperties properties,
                      @Value("${aios.trm-service.url}") URI endpoint,
                      @Value("${aios.trm-service.connect-timeout:5s}") Duration connectTimeout,
                      @Value("${aios.trm-service.read-timeout:10s}") Duration requestTimeout) {
        this(locator, properties,
                HttpClient.newBuilder().connectTimeout(connectTimeout).build(), endpoint, requestTimeout);
    }

    TrmService(InsumosLocator locator, AiosProperties properties, HttpClient httpClient,
               URI endpoint, Duration requestTimeout) {
        this.locator = locator;
        this.properties = properties;
        this.httpClient = httpClient;
        this.endpoint = endpoint;
        this.requestTimeout = requestTimeout;
    }

    public BigDecimal obtener(LocalDate fechaCorte) {
        return cache.computeIfAbsent(fechaCorte, this::consultarConContingencia);
    }

    private BigDecimal consultarConContingencia(LocalDate fechaCorte) {
        try {
            BigDecimal trm = consultarServicio(fechaCorte);
            log.info("TRM obtenida del servicio web de la Superfinanciera para fechaCorte={}: {}",
                    fechaCorte, trm);
            return trm;
        } catch (Exception e) {
            log.warn("Falló la consulta de TRM al servicio web para fechaCorte={}; se usará PIB_PEA_TRM_DG. Causa: {}",
                    fechaCorte, e.getMessage());
            BigDecimal trm = readFromFile(fechaCorte);
            log.info("TRM obtenida del archivo de contingencia para fechaCorte={}: {}", fechaCorte, trm);
            return trm;
        }
    }

    private BigDecimal consultarServicio(LocalDate fechaCorte) throws Exception {
        String body = """
                <soapenv:Envelope xmlns:soapenv="http://schemas.xmlsoap.org/soap/envelope/"
                    xmlns:act="%s">
                  <soapenv:Header/>
                  <soapenv:Body>
                    <act:queryTCRM>
                      <tcrmQueryAssociatedDate>%s</tcrmQueryAssociatedDate>
                    </act:queryTCRM>
                  </soapenv:Body>
                </soapenv:Envelope>
                """.formatted(SOAP_NAMESPACE, fechaCorte);

        HttpRequest request = HttpRequest.newBuilder(endpoint)
                .timeout(requestTimeout)
                .header("Content-Type", "text/xml; charset=UTF-8")
                .POST(HttpRequest.BodyPublishers.ofString(body, StandardCharsets.UTF_8))
                .build();
        HttpResponse<byte[]> response = httpClient.send(request, HttpResponse.BodyHandlers.ofByteArray());
        if (response.statusCode() < 200 || response.statusCode() >= 300) {
            throw new IllegalStateException("HTTP " + response.statusCode());
        }

        DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
        factory.setNamespaceAware(true);
        factory.setFeature("http://apache.org/xml/features/disallow-doctype-decl", true);
        factory.setFeature("http://xml.org/sax/features/external-general-entities", false);
        factory.setFeature("http://xml.org/sax/features/external-parameter-entities", false);
        factory.setXIncludeAware(false);
        factory.setExpandEntityReferences(false);
        var document = factory.newDocumentBuilder().parse(new ByteArrayInputStream(response.body()));
        var values = document.getElementsByTagNameNS("*", "value");
        if (values.getLength() == 0) {
            throw new IllegalStateException("La respuesta SOAP no contiene el campo value");
        }
        BigDecimal trm = new BigDecimal(values.item(0).getTextContent().trim());
        if (trm.signum() <= 0) {
            throw new IllegalStateException("El servicio devolvió una TRM no positiva");
        }
        return trm;
    }

    BigDecimal readFromFile(LocalDate fechaCorte) {
        try {
            var seriesFile = locator.findRequired("PIB_PEA_TRM_DG", fechaCorte);
            long maxBytes = (properties.maxPoiFileMb() == null ? 40L : properties.maxPoiFileMb()) * 1024L * 1024L;
            if (java.nio.file.Files.size(seriesFile) > maxBytes) {
                throw new IllegalStateException("PIB_PEA_TRM_DG excede el tamaño máximo permitido para POI");
            }
            try (Workbook wb = WorkbookFactory.create(seriesFile.toFile(), null, true)) {
                Sheet sheet = wb.getSheet("Hoja1");
                if (sheet == null) sheet = wb.getSheetAt(0);
                BigDecimal trm = null;
                LocalDate mejorFecha = LocalDate.MIN;
                for (Row row : sheet) {
                    LocalDate fecha = cellAsDate(row.getCell(1));
                    BigDecimal valor = numeric(row.getCell(2));
                    if (fecha != null && !fecha.isAfter(fechaCorte) && valor.signum() > 0 && !fecha.isBefore(mejorFecha)) {
                        mejorFecha = fecha;
                        trm = valor;
                    }
                }
                if (trm == null) throw new IllegalStateException("No hay una TRM válida en el archivo");
                return trm;
            }
        } catch (Exception e) {
            throw new IllegalStateException("No fue posible leer la TRM del archivo de contingencia", e);
        }
    }

    private LocalDate cellAsDate(Cell cell) {
        if (cell == null) return null;
        try {
            if (cell.getCellType() == CellType.NUMERIC) {
                return cell.getLocalDateTimeCellValue().toLocalDate();
            }
            return LocalDate.parse(cell.getStringCellValue().trim());
        } catch (Exception ignored) {
            return null;
        }
    }

    private BigDecimal numeric(Cell cell) {
        if (cell == null) return BigDecimal.ZERO;
        try {
            if (cell.getCellType() == CellType.NUMERIC || cell.getCellType() == CellType.FORMULA) {
                return BigDecimal.valueOf(cell.getNumericCellValue());
            }
            return new BigDecimal(cell.getStringCellValue().trim().replace(',', '.'));
        } catch (Exception ignored) {
            return BigDecimal.ZERO;
        }
    }
}
