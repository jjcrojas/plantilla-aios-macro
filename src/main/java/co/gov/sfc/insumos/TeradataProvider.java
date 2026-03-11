package co.gov.sfc.insumos;

import org.springframework.stereotype.Component;

import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Path;

@Component
public class TeradataProvider implements InsumosProvider {

    @Override
    public InputStream open(String fileName) throws IOException {
        throw new UnsupportedOperationException("TeradataProvider aún no implementado en Fase 1");
    }

    @Override
    public Path resolve(String fileName) {
        throw new UnsupportedOperationException("TeradataProvider aún no implementado en Fase 1");
    }
}
