package com.github.wnameless.workbook.yamlworkbook;

import static org.junit.jupiter.api.Assertions.*;
import java.io.IOException;
import java.io.InputStream;
import java.nio.charset.StandardCharsets;
import java.util.List;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.jupiter.api.Test;
import org.yaml.snakeyaml.nodes.MappingNode;
import org.yaml.snakeyaml.nodes.Node;
import org.yaml.snakeyaml.nodes.NodeTuple;
import org.yaml.snakeyaml.nodes.ScalarNode;
import org.yaml.snakeyaml.nodes.SequenceNode;

class YamlWorkbookReaderTest {

  @Test
  void testSimpleYamlRoundTrip() throws IOException {
    String yaml = loadYaml("yaml/simple.yaml");
    Workbook workbook = YamlWorkbook.toWorkbook(yaml);

    YamlWorkbookReader reader = YamlWorkbookReader.builder().build();
    List<Node> nodes = reader.fromWorkbook(workbook);

    assertEquals(1, nodes.size());
    assertTrue(nodes.get(0) instanceof MappingNode);

    MappingNode mapping = (MappingNode) nodes.get(0);
    List<NodeTuple> tuples = mapping.getValue();
    assertEquals(3, tuples.size());

    // name: John Doe
    assertEquals("name", getScalarValue(tuples.get(0).getKeyNode()));
    assertEquals("John Doe", getScalarValue(tuples.get(0).getValueNode()));

    // age: 30
    assertEquals("age", getScalarValue(tuples.get(1).getKeyNode()));
    assertEquals("30", getScalarValue(tuples.get(1).getValueNode()));

    // city: New York
    assertEquals("city", getScalarValue(tuples.get(2).getKeyNode()));
    assertEquals("New York", getScalarValue(tuples.get(2).getValueNode()));

    workbook.close();
  }

  @Test
  void testNestedYamlRoundTrip() throws IOException {
    String yaml = loadYaml("yaml/nested.yaml");
    Workbook workbook = YamlWorkbook.toWorkbook(yaml);

    YamlWorkbookReader reader = YamlWorkbookReader.builder().build();
    List<Node> nodes = reader.fromWorkbook(workbook);

    assertEquals(1, nodes.size());
    assertTrue(nodes.get(0) instanceof MappingNode);

    MappingNode root = (MappingNode) nodes.get(0);
    assertEquals(1, root.getValue().size());

    // person:
    NodeTuple personTuple = root.getValue().get(0);
    assertEquals("person", getScalarValue(personTuple.getKeyNode()));
    assertTrue(personTuple.getValueNode() instanceof MappingNode);

    MappingNode person = (MappingNode) personTuple.getValueNode();
    assertEquals(3, person.getValue().size()); // name, address, contact

    // name: Jane Smith
    assertEquals("name", getScalarValue(person.getValue().get(0).getKeyNode()));
    assertEquals("Jane Smith", getScalarValue(person.getValue().get(0).getValueNode()));

    // address:
    assertEquals("address", getScalarValue(person.getValue().get(1).getKeyNode()));
    assertTrue(person.getValue().get(1).getValueNode() instanceof MappingNode);

    MappingNode address = (MappingNode) person.getValue().get(1).getValueNode();
    assertEquals("street", getScalarValue(address.getValue().get(0).getKeyNode()));
    assertEquals("123 Main St", getScalarValue(address.getValue().get(0).getValueNode()));

    workbook.close();
  }

  @Test
  void testSequenceYamlRoundTrip() throws IOException {
    String yaml = loadYaml("yaml/sequence.yaml");
    Workbook workbook = YamlWorkbook.toWorkbook(yaml);

    YamlWorkbookReader reader = YamlWorkbookReader.builder().build();
    List<Node> nodes = reader.fromWorkbook(workbook);

    assertEquals(1, nodes.size());
    assertTrue(nodes.get(0) instanceof MappingNode);

    MappingNode root = (MappingNode) nodes.get(0);
    assertEquals(2, root.getValue().size()); // fruits and colors

    // fruits:
    NodeTuple fruitsTuple = root.getValue().get(0);
    assertEquals("fruits", getScalarValue(fruitsTuple.getKeyNode()));
    assertTrue(fruitsTuple.getValueNode() instanceof SequenceNode);

    SequenceNode fruits = (SequenceNode) fruitsTuple.getValueNode();
    assertEquals(3, fruits.getValue().size());
    assertEquals("apple", getScalarValue(fruits.getValue().get(0)));
    assertEquals("banana", getScalarValue(fruits.getValue().get(1)));
    assertEquals("cherry", getScalarValue(fruits.getValue().get(2)));

    // colors:
    NodeTuple colorsTuple = root.getValue().get(1);
    assertEquals("colors", getScalarValue(colorsTuple.getKeyNode()));
    assertTrue(colorsTuple.getValueNode() instanceof SequenceNode);

    SequenceNode colors = (SequenceNode) colorsTuple.getValueNode();
    assertEquals(3, colors.getValue().size());
    assertEquals("red", getScalarValue(colors.getValue().get(0)));
    assertEquals("green", getScalarValue(colors.getValue().get(1)));
    assertEquals("blue", getScalarValue(colors.getValue().get(2)));

    workbook.close();
  }

  @Test
  void testScalarOnlyRoundTrip() throws IOException {
    Workbook workbook = YamlWorkbook.toWorkbook("hello world");

    YamlWorkbookReader reader = YamlWorkbookReader.builder().build();
    List<Node> nodes = reader.fromWorkbook(workbook);

    assertEquals(1, nodes.size());
    assertTrue(nodes.get(0) instanceof ScalarNode);
    assertEquals("hello world", getScalarValue(nodes.get(0)));

    workbook.close();
  }

  @Test
  void testMultiDocumentRoundTrip() throws IOException {
    String yaml = loadYaml("yaml/multidoc.yaml");
    Workbook workbook = YamlWorkbook.toWorkbook(yaml);

    YamlWorkbookReader reader = YamlWorkbookReader.builder().build();
    List<Node> nodes = reader.fromWorkbook(workbook);

    assertEquals(3, nodes.size());

    // Document One
    assertTrue(nodes.get(0) instanceof MappingNode);
    MappingNode doc1 = (MappingNode) nodes.get(0);
    assertEquals("name", getScalarValue(doc1.getValue().get(0).getKeyNode()));
    assertEquals("Document One", getScalarValue(doc1.getValue().get(0).getValueNode()));

    // Document Two
    assertTrue(nodes.get(1) instanceof MappingNode);
    MappingNode doc2 = (MappingNode) nodes.get(1);
    assertEquals("name", getScalarValue(doc2.getValue().get(0).getKeyNode()));
    assertEquals("Document Two", getScalarValue(doc2.getValue().get(0).getValueNode()));

    // Document Three
    assertTrue(nodes.get(2) instanceof MappingNode);
    MappingNode doc3 = (MappingNode) nodes.get(2);
    assertEquals("name", getScalarValue(doc3.getValue().get(0).getKeyNode()));
    assertEquals("Document Three", getScalarValue(doc3.getValue().get(0).getValueNode()));

    workbook.close();
  }

  @Test
  void testEmptyWorkbook() {
    YamlWorkbookReader reader = YamlWorkbookReader.builder().build();
    List<Node> nodes = reader.fromWorkbook(null);

    assertTrue(nodes.isEmpty());
  }

  @Test
  void testCommentsPreserved() throws IOException {
    String yaml = loadYaml("yaml/comments.yaml");
    Workbook workbook = YamlWorkbook.toWorkbook(yaml);

    YamlWorkbookReader reader = YamlWorkbookReader.builder().build();
    List<Node> nodes = reader.fromWorkbook(workbook);

    assertEquals(1, nodes.size());
    assertTrue(nodes.get(0) instanceof MappingNode);

    MappingNode root = (MappingNode) nodes.get(0);
    // Check that block comments are preserved
    assertNotNull(root.getValue().get(0).getKeyNode().getBlockComments());
    assertFalse(root.getValue().get(0).getKeyNode().getBlockComments().isEmpty());

    workbook.close();
  }

  @Test
  void testInlineCommentsPreserved() throws IOException {
    String yaml = loadYaml("yaml/comments.yaml");
    Workbook workbook = YamlWorkbook.toWorkbook(yaml);

    YamlWorkbookReader reader = YamlWorkbookReader.builder().build();
    List<Node> nodes = reader.fromWorkbook(workbook);

    MappingNode root = (MappingNode) nodes.get(0);
    MappingNode database = (MappingNode) root.getValue().get(0).getValueNode();

    // host: localhost # primary host
    Node hostValue = database.getValue().get(0).getValueNode();
    assertNotNull(hostValue.getInLineComments());
    assertFalse(hostValue.getInLineComments().isEmpty());

    workbook.close();
  }

  @Test
  void testComplexYamlRoundTrip() throws IOException {
    String yaml = loadYaml("yaml/complex.yaml");
    Workbook workbook = YamlWorkbook.toWorkbook(yaml);

    YamlWorkbookReader reader = YamlWorkbookReader.builder().build();
    List<Node> nodes = reader.fromWorkbook(workbook);

    assertFalse(nodes.isEmpty());
    assertTrue(nodes.get(0) instanceof MappingNode);

    workbook.close();
  }

  private String loadYaml(String resourcePath) throws IOException {
    try (InputStream is = getClass().getClassLoader().getResourceAsStream(resourcePath)) {
      if (is == null) {
        throw new IOException("Resource not found: " + resourcePath);
      }
      return new String(is.readAllBytes(), StandardCharsets.UTF_8);
    }
  }

  private String getScalarValue(Node node) {
    if (node instanceof ScalarNode) {
      return ((ScalarNode) node).getValue();
    }
    return null;
  }

}
