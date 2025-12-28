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

/**
 * Static utility class providing convenience methods for bidirectional conversion between YAML and
 * Excel workbooks.
 * <p>
 * This class offers simple entry points for common operations:
 * <ul>
 * <li>{@link #toWorkbook(String, String...)} - Convert YAML strings to Excel workbook</li>
 * <li>{@link #fromWorkbook(Workbook)} - Convert Excel workbook to SnakeYAML Node list</li>
 * <li>{@link #toYaml(Workbook)} - Convert Excel workbook to YAML string</li>
 * <li>{@link #writerBuilder()} / {@link #readerBuilder()} - Access builder APIs for
 * customization</li>
 * </ul>
 *
 * @author Wei-Ming Wu
 */
@UtilityClass
public class YamlWorkbook {

  /**
   * Returns a new writer builder for converting YAML to Excel.
   *
   * @return a new {@link YamlWorkbookWriter.YamlWorkbookWriterBuilder}
   */
  public YamlWorkbookWriter.YamlWorkbookWriterBuilder writerBuilder() {
    return YamlWorkbookWriter.builder();
  }

  /**
   * Returns a new reader builder for converting Excel to YAML.
   *
   * @return a new {@link YamlWorkbookReader.YamlWorkbookReaderBuilder}
   */
  public YamlWorkbookReader.YamlWorkbookReaderBuilder readerBuilder() {
    return YamlWorkbookReader.builder();
  }

  /**
   * Returns a writer builder pre-configured for PREFIX indentation mode.
   *
   * @return a new {@link YamlWorkbookWriter.YamlWorkbookWriterBuilder} with PREFIX mode
   */
  public YamlWorkbookWriter.YamlWorkbookWriterBuilder prefixWriterBuilder() {
    return YamlWorkbookWriter.builder().indentationMode(IndentationMode.PREFIX);
  }

  /**
   * Returns a reader builder pre-configured for PREFIX indentation mode.
   *
   * @return a new {@link YamlWorkbookReader.YamlWorkbookReaderBuilder} with PREFIX mode
   */
  public YamlWorkbookReader.YamlWorkbookReaderBuilder prefixReaderBuilder() {
    return YamlWorkbookReader.builder().indentationMode(IndentationMode.PREFIX);
  }

  /**
   * Converts YAML content strings to an Excel workbook.
   *
   * @param yamlContent the primary YAML content
   * @param yamlContents additional YAML content strings (optional)
   * @return the generated Excel workbook
   */
  public Workbook toWorkbook(String yamlContent, String... yamlContents) {
    int length = yamlContents == null ? 0 : yamlContents.length;
    StringReader[] yamlContentReaders = new StringReader[length];
    for (int i = 0; i < length; i++) {
      yamlContentReaders[i] = new StringReader(yamlContents[i]);
    }
    return YamlWorkbookWriter.builder().build().toWorkbook(new StringReader(yamlContent),
        yamlContentReaders);
  }

  /**
   * Converts YAML content from StringReaders to an Excel workbook.
   *
   * @param yamlContent the primary YAML content reader
   * @param yamlContents additional YAML content readers (optional)
   * @return the generated Excel workbook
   */
  public Workbook toWorkbook(StringReader yamlContent, StringReader... yamlContents) {
    if (yamlContents == null) {
      yamlContents = new StringReader[0];
    }
    return YamlWorkbookWriter.builder().build().toWorkbook(yamlContent, yamlContents);
  }

  /**
   * Converts YAML content from InputStreams to an Excel workbook using UTF-8 charset.
   *
   * @param yamlContent the primary YAML content input stream
   * @param yamlContents additional YAML content input streams (optional)
   * @return the generated Excel workbook
   */
  public Workbook toWorkbook(InputStream yamlContent, InputStream... yamlContents) {
    return toWorkbook(yamlContent, StandardCharsets.UTF_8, yamlContents);
  }

  /**
   * Converts YAML content from InputStreams to an Excel workbook using specified charset.
   *
   * @param yamlContent the primary YAML content input stream
   * @param charset the character set for reading input streams
   * @param yamlContents additional YAML content input streams (optional)
   * @return the generated Excel workbook
   */
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

  /**
   * Converts an Excel workbook to a list of SnakeYAML Node objects.
   *
   * @param workbook the Excel workbook to convert
   * @return a list of YAML document nodes
   */
  public List<Node> fromWorkbook(Workbook workbook) {
    return YamlWorkbookReader.builder().build().fromWorkbook(workbook);
  }

  /**
   * Converts an Excel workbook to a YAML string.
   *
   * @param workbook the Excel workbook to convert
   * @return the YAML string representation
   */
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

}
