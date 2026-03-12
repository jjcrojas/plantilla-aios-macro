package co.gov.sfc.insumos;

import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Path;

public interface InsumosProvider {
    InputStream open(String fileName) throws IOException;
    Path resolve(String fileName);
}
