package lu.nyo.excel.renderer;

import org.apache.commons.io.IOUtils;

import java.io.IOException;
import java.io.OutputStream;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Objects;

import static java.nio.file.Files.newOutputStream;
import static java.nio.file.Path.of;

public class Utils {
    private static final String EXPORT_PATH = System.getProperty("user.dir");

    public static String getTestCss() throws IOException {
        return IOUtils.toString(Objects.requireNonNull(ExcelFileRenderer.class.getClassLoader().getResourceAsStream("test.css")), StandardCharsets.UTF_8);
    }

    public static OutputStream getOutputStream(String fileName) throws IOException {
        Path generated = of(EXPORT_PATH, "generated");
        if (!Files.exists(generated)) {
            Files.createDirectory(generated);
        }
        return newOutputStream(of(EXPORT_PATH, "generated", fileName + ".xlsx"));
    }
}
