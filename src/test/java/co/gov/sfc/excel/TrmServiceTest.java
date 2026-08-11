package co.gov.sfc.excel;

import co.gov.sfc.config.AiosProperties;
import co.gov.sfc.insumos.InsumosLocator;
import org.junit.jupiter.api.Test;

import java.math.BigDecimal;
import java.net.URI;
import java.net.http.HttpClient;
import java.net.http.HttpRequest;
import java.net.http.HttpResponse;
import java.nio.file.Path;
import java.time.Duration;
import java.time.LocalDate;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.mockito.ArgumentMatchers.any;
import static org.mockito.Mockito.*;

class TrmServiceTest {

    private static final LocalDate FECHA = LocalDate.of(2025, 6, 30);

    @Test
    @SuppressWarnings({"unchecked", "rawtypes"})
    void shouldQuerySoapOnlyOnceAndCacheValueByDate() throws Exception {
        HttpClient client = mock(HttpClient.class);
        HttpResponse<byte[]> response = mock(HttpResponse.class);
        when(response.statusCode()).thenReturn(200);
        when(response.body()).thenReturn("""
                <soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
                  <soap:Body><queryTCRMResponse><return><value>4192.03</value></return></queryTCRMResponse></soap:Body>
                </soap:Envelope>
                """.getBytes());
        when(client.send(any(HttpRequest.class), any(HttpResponse.BodyHandler.class))).thenReturn((HttpResponse) response);
        TrmService service = service(client);

        assertEquals(new BigDecimal("4192.03"), service.obtener(FECHA));
        assertEquals(new BigDecimal("4192.03"), service.obtener(FECHA));

        verify(client, times(1)).send(any(HttpRequest.class), any(HttpResponse.BodyHandler.class));
    }

    @Test
    @SuppressWarnings({"unchecked", "rawtypes"})
    void shouldUseAndCacheFileValueWhenServiceFails() throws Exception {
        HttpClient client = mock(HttpClient.class);
        HttpResponse<byte[]> response = mock(HttpResponse.class);
        when(response.statusCode()).thenReturn(503);
        when(client.send(any(HttpRequest.class), any(HttpResponse.BodyHandler.class))).thenReturn((HttpResponse) response);
        TrmService service = spy(service(client));
        doReturn(new BigDecimal("4150.75")).when(service).readFromFile(FECHA);

        assertEquals(new BigDecimal("4150.75"), service.obtener(FECHA));
        assertEquals(new BigDecimal("4150.75"), service.obtener(FECHA));

        verify(client, times(1)).send(any(HttpRequest.class), any(HttpResponse.BodyHandler.class));
        verify(service, times(1)).readFromFile(FECHA);
    }

    private TrmService service(HttpClient client) {
        AiosProperties properties = new AiosProperties(Path.of("."), Path.of("."), Path.of("."), 40, false);
        return new TrmService(mock(InsumosLocator.class), properties, client,
                URI.create("https://example.test/trm"), Duration.ofSeconds(2));
    }
}
