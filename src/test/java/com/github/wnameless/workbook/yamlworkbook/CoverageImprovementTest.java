package com.github.wnameless.workbook.yamlworkbook;

import static org.junit.jupiter.api.Assertions.*;
import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.StringReader;
import java.nio.charset.StandardCharsets;
import java.util.List;
import java.util.Map;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.jupiter.api.Test;
import org.yaml.snakeyaml.nodes.MappingNode;
import org.yaml.snakeyaml.nodes.Node;
import org.yaml.snakeyaml.nodes.ScalarNode;
import org.yaml.snakeyaml.nodes.SequenceNode;
import org.yaml.snakeyaml.nodes.Tag;
import com.github.wnameless.json.jsonschemadatagenerator.ObjectMapperFactory;
import tools.jackson.databind.JsonNode;
import tools.jackson.databind.node.ArrayNode;
import tools.jackson.databind.node.NullNode;
import tools.jackson.databind.node.ObjectNode;

/**
 * Additional tests to improve code coverage for:
 * - CommentType enum
 * - YamlWorkbook utility class
 * - JsonNodeToYamlNodeConverter utility class
 */
class CoverageImprovementTest {

  // ==================== CommentType Tests ====================

  @Test
  void testCommentTypeEnumValues() {
    // Test that all enum values are accessible
    CommentType[] values = CommentType.values();
    assertEquals(7, values.length);

    // Verify each value exists
    assertEquals(CommentType.DOCUMENT, CommentType.valueOf("DOCUMENT"));
    assertEquals(CommentType.OBJECT, CommentType.valueOf("OBJECT"));
    assertEquals(CommentType.ARRAY, CommentType.valueOf("ARRAY"));
    assertEquals(CommentType.KEY, CommentType.valueOf("KEY"));
    assertEquals(CommentType.VALUE, CommentType.valueOf("VALUE"));
    assertEquals(CommentType.KEY_VALUE_PAIR, CommentType.valueOf("KEY_VALUE_PAIR"));
    assertEquals(CommentType.ITEM, CommentType.valueOf("ITEM"));
  }

  // ==================== YamlWorkbook Tests ====================

  @Test
  void testPrefixWriterBuilder() throws IOException {
    // Test that prefixWriterBuilder returns a builder configured for PREFIX mode
    YamlWorkbookWriter.YamlWorkbookWriterBuilder builder = YamlWorkbook.prefixWriterBuilder();
    assertNotNull(builder);

    String yaml = """
        name: John
        address:
          city: NYC
        """;

    Workbook workbook = builder.build().toWorkbook(new StringReader(yaml));
    assertNotNull(workbook);
    assertEquals(1, workbook.getNumberOfSheets());
    workbook.close();
  }

  @Test
  void testPrefixReaderBuilder() throws IOException {
    // Test that prefixReaderBuilder returns a builder configured for PREFIX mode
    YamlWorkbookReader.YamlWorkbookReaderBuilder builder = YamlWorkbook.prefixReaderBuilder();
    assertNotNull(builder);

    // Create a workbook using PREFIX mode writer
    String yaml = """
        name: Test
        value: 123
        """;
    Workbook workbook = YamlWorkbook.prefixWriterBuilder().build().toWorkbook(new StringReader(yaml));

    // Read it back using PREFIX mode reader
    List<Node> nodes = builder.build().fromWorkbook(workbook);
    assertNotNull(nodes);
    assertFalse(nodes.isEmpty());
    workbook.close();
  }

  @Test
  void testToWorkbookWithInputStream() throws IOException {
    // Test toWorkbook(InputStream, InputStream...)
    String yaml = "name: John\nage: 30";
    InputStream is = new ByteArrayInputStream(yaml.getBytes(StandardCharsets.UTF_8));

    Workbook workbook = YamlWorkbook.toWorkbook(is);
    assertNotNull(workbook);
    assertEquals(1, workbook.getNumberOfSheets());
    workbook.close();
  }

  @Test
  void testToWorkbookWithInputStreamAndCharset() throws IOException {
    // Test toWorkbook(InputStream, Charset, InputStream...)
    String yaml = "name: John\nage: 30";
    InputStream is = new ByteArrayInputStream(yaml.getBytes(StandardCharsets.UTF_8));

    Workbook workbook = YamlWorkbook.toWorkbook(is, StandardCharsets.UTF_8);
    assertNotNull(workbook);
    assertEquals(1, workbook.getNumberOfSheets());
    workbook.close();
  }

  @Test
  void testToWorkbookWithMultipleInputStreams() throws IOException {
    // Test toWorkbook with multiple InputStreams
    String yaml1 = "doc: one";
    String yaml2 = "doc: two";
    InputStream is1 = new ByteArrayInputStream(yaml1.getBytes(StandardCharsets.UTF_8));
    InputStream is2 = new ByteArrayInputStream(yaml2.getBytes(StandardCharsets.UTF_8));

    Workbook workbook = YamlWorkbook.toWorkbook(is1, is2);
    assertNotNull(workbook);
    workbook.close();
  }

  @Test
  void testToWorkbookWithMultipleInputStreamsAndCharset() throws IOException {
    // Test toWorkbook with multiple InputStreams and explicit charset
    String yaml1 = "doc: one";
    String yaml2 = "doc: two";
    InputStream is1 = new ByteArrayInputStream(yaml1.getBytes(StandardCharsets.UTF_8));
    InputStream is2 = new ByteArrayInputStream(yaml2.getBytes(StandardCharsets.UTF_8));

    Workbook workbook = YamlWorkbook.toWorkbook(is1, StandardCharsets.UTF_8, is2);
    assertNotNull(workbook);
    workbook.close();
  }

  @Test
  void testToWorkbookWithNullVarargs() throws IOException {
    // Test toWorkbook with null varargs
    String yaml = "name: John";

    Workbook workbook = YamlWorkbook.toWorkbook(yaml, (String[]) null);
    assertNotNull(workbook);
    workbook.close();
  }

  @Test
  void testToWorkbookWithStringReaderNullVarargs() throws IOException {
    // Test toWorkbook(StringReader, StringReader...) with null varargs
    String yaml = "name: John";
    StringReader reader = new StringReader(yaml);

    Workbook workbook = YamlWorkbook.toWorkbook(reader, (StringReader[]) null);
    assertNotNull(workbook);
    workbook.close();
  }

  @Test
  void testToWorkbookWithInputStreamNullVarargs() throws IOException {
    // Test toWorkbook(InputStream, Charset, InputStream...) with null varargs
    String yaml = "name: John";
    InputStream is = new ByteArrayInputStream(yaml.getBytes(StandardCharsets.UTF_8));

    Workbook workbook = YamlWorkbook.toWorkbook(is, StandardCharsets.UTF_8, (InputStream[]) null);
    assertNotNull(workbook);
    workbook.close();
  }

  @Test
  void testToYamlWithContent() throws IOException {
    // Test toYaml with a workbook containing YAML content
    String yaml = """
        name: John Doe
        age: 30
        """;
    Workbook workbook = YamlWorkbook.toWorkbook(yaml);

    String result = YamlWorkbook.toYaml(workbook);
    assertNotNull(result);
    assertFalse(result.isEmpty());
    assertTrue(result.contains("name"));
    assertTrue(result.contains("John Doe"));

    workbook.close();
  }

  @Test
  void testToYamlWithEmptyWorkbook() throws IOException {
    // Test toYaml with empty workbook
    Workbook workbook = YamlWorkbook.toWorkbook("");
    String result = YamlWorkbook.toYaml(workbook);
    assertEquals("", result);
    workbook.close();
  }

  @Test
  void testWriterBuilder() {
    // Test writerBuilder returns valid builder
    var builder = YamlWorkbook.writerBuilder();
    assertNotNull(builder);
  }

  @Test
  void testReaderBuilder() {
    // Test readerBuilder returns valid builder
    var builder = YamlWorkbook.readerBuilder();
    assertNotNull(builder);
  }

  // ==================== JsonNodeToYamlNodeConverter Tests ====================

  @Test
  void testConvertNull() {
    // Test convert with null input
    Node result = JsonNodeToYamlNodeConverter.convert(null);
    assertNotNull(result);
    assertTrue(result instanceof ScalarNode);
    assertEquals(Tag.NULL, result.getTag());
    assertEquals("null", ((ScalarNode) result).getValue());
  }

  @Test
  void testConvertNullNode() {
    // Test convert with NullNode
    NullNode nullNode = NullNode.instance;
    Node result = JsonNodeToYamlNodeConverter.convert(nullNode);
    assertNotNull(result);
    assertTrue(result instanceof ScalarNode);
    assertEquals(Tag.NULL, result.getTag());
  }

  @Test
  void testConvertEmptyObjectNode() {
    // Test convert with empty ObjectNode
    ObjectNode objectNode = ObjectMapperFactory.getObjectMapper().createObjectNode();
    Node result = JsonNodeToYamlNodeConverter.convert(objectNode);
    assertNotNull(result);
    assertTrue(result instanceof MappingNode);
    MappingNode mappingNode = (MappingNode) result;
    assertTrue(mappingNode.getValue().isEmpty());
  }

  @Test
  void testConvertObjectNodeWithProperties() {
    // Test convert with ObjectNode containing properties
    ObjectNode objectNode = ObjectMapperFactory.getObjectMapper().createObjectNode();
    objectNode.put("name", "John");
    objectNode.put("age", 30);
    objectNode.put("active", true);

    Node result = JsonNodeToYamlNodeConverter.convert(objectNode);
    assertNotNull(result);
    assertTrue(result instanceof MappingNode);
    MappingNode mappingNode = (MappingNode) result;
    assertEquals(3, mappingNode.getValue().size());
  }

  @Test
  void testConvertEmptyArrayNode() {
    // Test convert with empty ArrayNode
    ArrayNode arrayNode = ObjectMapperFactory.getObjectMapper().createArrayNode();
    Node result = JsonNodeToYamlNodeConverter.convert(arrayNode);
    assertNotNull(result);
    assertTrue(result instanceof SequenceNode);
    SequenceNode sequenceNode = (SequenceNode) result;
    assertTrue(sequenceNode.getValue().isEmpty());
  }

  @Test
  void testConvertArrayNodeWithElements() {
    // Test convert with ArrayNode containing elements
    ArrayNode arrayNode = ObjectMapperFactory.getObjectMapper().createArrayNode();
    arrayNode.add("apple");
    arrayNode.add("banana");
    arrayNode.add(123);

    Node result = JsonNodeToYamlNodeConverter.convert(arrayNode);
    assertNotNull(result);
    assertTrue(result instanceof SequenceNode);
    SequenceNode sequenceNode = (SequenceNode) result;
    assertEquals(3, sequenceNode.getValue().size());
  }

  @Test
  void testConvertStringNode() {
    // Test convert with string value
    JsonNode stringNode = ObjectMapperFactory.getObjectMapper().valueToTree("hello");
    Node result = JsonNodeToYamlNodeConverter.convert(stringNode);
    assertNotNull(result);
    assertTrue(result instanceof ScalarNode);
    assertEquals(Tag.STR, result.getTag());
    assertEquals("hello", ((ScalarNode) result).getValue());
  }

  @Test
  void testConvertIntegerNode() {
    // Test convert with integer value
    JsonNode intNode = ObjectMapperFactory.getObjectMapper().valueToTree(42);
    Node result = JsonNodeToYamlNodeConverter.convert(intNode);
    assertNotNull(result);
    assertTrue(result instanceof ScalarNode);
    assertEquals(Tag.INT, result.getTag());
    assertEquals("42", ((ScalarNode) result).getValue());
  }

  @Test
  void testConvertLongNode() {
    // Test convert with long value
    JsonNode longNode = ObjectMapperFactory.getObjectMapper().valueToTree(9876543210L);
    Node result = JsonNodeToYamlNodeConverter.convert(longNode);
    assertNotNull(result);
    assertTrue(result instanceof ScalarNode);
    assertEquals(Tag.INT, result.getTag());
  }

  @Test
  void testConvertFloatNode() {
    // Test convert with float/double value
    JsonNode floatNode = ObjectMapperFactory.getObjectMapper().valueToTree(3.14);
    Node result = JsonNodeToYamlNodeConverter.convert(floatNode);
    assertNotNull(result);
    assertTrue(result instanceof ScalarNode);
    assertEquals(Tag.FLOAT, result.getTag());
  }

  @Test
  void testConvertBooleanTrueNode() {
    // Test convert with boolean true
    JsonNode boolNode = ObjectMapperFactory.getObjectMapper().valueToTree(true);
    Node result = JsonNodeToYamlNodeConverter.convert(boolNode);
    assertNotNull(result);
    assertTrue(result instanceof ScalarNode);
    assertEquals(Tag.BOOL, result.getTag());
    assertEquals("true", ((ScalarNode) result).getValue());
  }

  @Test
  void testConvertBooleanFalseNode() {
    // Test convert with boolean false
    JsonNode boolNode = ObjectMapperFactory.getObjectMapper().valueToTree(false);
    Node result = JsonNodeToYamlNodeConverter.convert(boolNode);
    assertNotNull(result);
    assertTrue(result instanceof ScalarNode);
    assertEquals(Tag.BOOL, result.getTag());
    assertEquals("false", ((ScalarNode) result).getValue());
  }

  @Test
  void testConvertNestedStructure() {
    // Test convert with nested object/array structure
    ObjectNode root = ObjectMapperFactory.getObjectMapper().createObjectNode();
    ObjectNode nested = ObjectMapperFactory.getObjectMapper().createObjectNode();
    ArrayNode array = ObjectMapperFactory.getObjectMapper().createArrayNode();

    array.add("item1");
    array.add("item2");
    nested.put("key", "value");
    nested.set("list", array);
    root.set("nested", nested);

    Node result = JsonNodeToYamlNodeConverter.convert(root);
    assertNotNull(result);
    assertTrue(result instanceof MappingNode);

    MappingNode mappingNode = (MappingNode) result;
    assertEquals(1, mappingNode.getValue().size());
  }

  @Test
  void testToMap() {
    // Test toMap method
    ObjectNode objectNode = ObjectMapperFactory.getObjectMapper().createObjectNode();
    objectNode.put("name", "John");
    objectNode.put("age", 30);

    Map<String, Object> result = JsonNodeToYamlNodeConverter.toMap(objectNode);
    assertNotNull(result);
    assertEquals(2, result.size());
    assertEquals("John", result.get("name"));
    assertEquals(30, result.get("age"));
  }

  @Test
  void testToMapWithNestedObject() {
    // Test toMap with nested structure
    ObjectNode root = ObjectMapperFactory.getObjectMapper().createObjectNode();
    ObjectNode nested = ObjectMapperFactory.getObjectMapper().createObjectNode();
    nested.put("city", "NYC");
    root.set("address", nested);
    root.put("name", "John");

    Map<String, Object> result = JsonNodeToYamlNodeConverter.toMap(root);
    assertNotNull(result);
    assertEquals(2, result.size());
    assertTrue(result.get("address") instanceof Map);
  }

}
