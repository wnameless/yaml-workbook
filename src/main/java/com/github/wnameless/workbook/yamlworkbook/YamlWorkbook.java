package com.github.wnameless.workbook.yamlworkbook;

import java.io.InputStream;
import java.io.InputStreamReader;
import java.io.StringReader;
import java.io.StringWriter;
import java.nio.charset.Charset;
import java.nio.charset.StandardCharsets;
import java.util.List;
import org.apache.poi.ss.usermodel.Workbook;
import org.yaml.snakeyaml.DumperOptions;
import org.yaml.snakeyaml.emitter.Emitter;
import org.yaml.snakeyaml.nodes.Node;
import org.yaml.snakeyaml.resolver.Resolver;
import org.yaml.snakeyaml.serializer.Serializer;
import lombok.experimental.UtilityClass;

@UtilityClass
public class YamlWorkbook {

  public YamlWorkbookWriter.YamlWorkbookWriterBuilder writerBuilder() {
    return YamlWorkbookWriter.builder();
  }

  public YamlWorkbookReader.YamlWorkbookReaderBuilder readerBuilder() {
    return YamlWorkbookReader.builder();
  }

  /**
   * Returns a writer builder pre-configured for PREFIX indentation mode.
   */
  public YamlWorkbookWriter.YamlWorkbookWriterBuilder prefixWriterBuilder() {
    return YamlWorkbookWriter.builder().indentationMode(IndentationMode.PREFIX);
  }

  /**
   * Returns a reader builder pre-configured for PREFIX indentation mode.
   */
  public YamlWorkbookReader.YamlWorkbookReaderBuilder prefixReaderBuilder() {
    return YamlWorkbookReader.builder().indentationMode(IndentationMode.PREFIX);
  }

  public Workbook toWorkbook(String yamlContent, String... yamlContents) {
    int length = yamlContents == null ? 0 : yamlContents.length;
    StringReader[] yamlContentReaders = new StringReader[length];
    for (int i = 0; i < length; i++) {
      yamlContentReaders[i] = new StringReader(yamlContents[i]);
    }
    return YamlWorkbookWriter.builder().build().toWorkbook(new StringReader(yamlContent),
        yamlContentReaders);
  }

  public Workbook toWorkbook(StringReader yamlContent, StringReader... yamlContents) {
    if (yamlContents == null) {
      yamlContents = new StringReader[0];
    }
    return YamlWorkbookWriter.builder().build().toWorkbook(yamlContent, yamlContents);
  }

  public Workbook toWorkbook(InputStream yamlContent, InputStream... yamlContents) {
    return toWorkbook(yamlContent, StandardCharsets.UTF_8, yamlContents);
  }

  public Workbook toWorkbook(InputStream yamlContent, Charset charset,
      InputStream... yamlContents) {
    int length = yamlContents == null ? 0 : yamlContents.length;
    InputStreamReader[] yamlContentReaders = new InputStreamReader[length];
    for (int i = 0; i < length; i++) {
      yamlContentReaders[i] = new InputStreamReader(yamlContents[i], charset);
    }
    return YamlWorkbookWriter.builder().build()
        .toWorkbook(new InputStreamReader(yamlContent, charset), yamlContentReaders);
  }

  public List<Node> fromWorkbook(Workbook workbook) {
    return YamlWorkbookReader.builder().build().fromWorkbook(workbook);
  }

  public String toYaml(Workbook workbook) {
    List<Node> nodes = fromWorkbook(workbook);
    if (nodes.isEmpty()) {
      return "";
    }

    StringWriter writer = new StringWriter();
    DumperOptions options = new DumperOptions();
    options.setDefaultFlowStyle(DumperOptions.FlowStyle.BLOCK);
    options.setProcessComments(true);

    Serializer serializer =
        new Serializer(new Emitter(writer, options), new Resolver(), options, null);
    try {
      serializer.open();
      for (Node node : nodes) {
        serializer.serialize(node);
      }
      serializer.close();
    } catch (Exception e) {
      throw new RuntimeException("Failed to serialize YAML nodes", e);
    }

    return writer.toString();
  }

  // public static void main(String[] args) throws IOException {
  // String jsonSchemaPath = "schema/ConfigurableNotification.schema.json";
  // try (InputStream is = YamlWorkbook.class.getClassLoader().getResourceAsStream(jsonSchemaPath))
  // {
  // if (is == null) {
  // throw new IOException("Resource not found: " + jsonSchemaPath);
  // }
  // var jsonSchema = new String(is.readAllBytes(), StandardCharsets.UTF_8);
  // var workbook = YamlWorkbook.writerBuilder().jsonSchema(jsonSchema)
  // .printMode(PrintMode.DATA_COLLECT).build().toWorkbook();

  // Path filePath = Paths.get("target/test-excel").resolve("main.xlsx");
  // try (FileOutputStream fos = new FileOutputStream(filePath.toFile())) {
  // workbook.write(fos);
  // }
  // }
  // }

}
