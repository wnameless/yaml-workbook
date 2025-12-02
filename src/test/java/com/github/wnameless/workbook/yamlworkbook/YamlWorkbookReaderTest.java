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
    // Use a simple YAML with only value inline comments (no key inline comments)
    // to test reader's inline comment preservation
    String yaml = """
        server:
          host: localhost # primary host
          port: 5432 # default port
        """;
    Workbook workbook = YamlWorkbook.toWorkbook(yaml);

    YamlWorkbookReader reader = YamlWorkbookReader.builder().build();
    List<Node> nodes = reader.fromWorkbook(workbook);

    MappingNode root = (MappingNode) nodes.get(0);
    MappingNode server = (MappingNode) root.getValue().get(0).getValueNode();

    // host: localhost # primary host
    Node hostValue = server.getValue().get(0).getValueNode();
    assertNotNull(hostValue.getInLineComments());
    assertFalse(hostValue.getInLineComments().isEmpty());

    // port: 5432 # default port
    Node portValue = server.getValue().get(1).getValueNode();
    assertNotNull(portValue.getInLineComments());
    assertFalse(portValue.getInLineComments().isEmpty());

    workbook.close();
  }

  @Test
  void testCommentsYamlRoundTrip() throws IOException {
    // Test round-trip for comments.yaml with the enhanced comment structure
    // including key inline comments (comments after keys)
    String yaml = loadYaml("yaml/comments.yaml");
    Workbook workbook = YamlWorkbook.toWorkbook(yaml);

    YamlWorkbookReader reader = YamlWorkbookReader.builder().build();
    List<Node> nodes = reader.fromWorkbook(workbook);

    assertEquals(1, nodes.size());
    assertTrue(nodes.get(0) instanceof MappingNode);

    MappingNode root = (MappingNode) nodes.get(0);
    assertEquals(1, root.getValue().size()); // database

    // database: # main database config
    NodeTuple databaseTuple = root.getValue().get(0);
    assertEquals("database", getScalarValue(databaseTuple.getKeyNode()));
    // Block comment preserved: "# This is a configuration file"
    assertNotNull(databaseTuple.getKeyNode().getBlockComments());
    assertFalse(databaseTuple.getKeyNode().getBlockComments().isEmpty());
    // Key inline comment preserved: "# main database config"
    assertNotNull(databaseTuple.getKeyNode().getInLineComments());
    assertFalse(databaseTuple.getKeyNode().getInLineComments().isEmpty());
    assertTrue(databaseTuple.getValueNode() instanceof MappingNode);

    MappingNode database = (MappingNode) databaseTuple.getValueNode();
    assertEquals(7, database.getValue().size()); // host, port, username, password, connection, replicas, allowed_ips

    // host: # server address localhost # primary host
    NodeTuple hostTuple = database.getValue().get(0);
    assertEquals("host", getScalarValue(hostTuple.getKeyNode()));
    // Key inline comment: "# server address"
    assertNotNull(hostTuple.getKeyNode().getInLineComments());
    assertFalse(hostTuple.getKeyNode().getInLineComments().isEmpty());
    assertEquals("localhost", getScalarValue(hostTuple.getValueNode()));
    // Value inline comment: "# primary host"
    assertNotNull(hostTuple.getValueNode().getInLineComments());
    assertFalse(hostTuple.getValueNode().getInLineComments().isEmpty());

    // port: 5432 # default PostgreSQL port
    NodeTuple portTuple = database.getValue().get(1);
    assertEquals("port", getScalarValue(portTuple.getKeyNode()));
    assertEquals("5432", getScalarValue(portTuple.getValueNode()));
    assertNotNull(portTuple.getValueNode().getInLineComments());

    // username: admin (no inline comment)
    NodeTuple usernameTuple = database.getValue().get(2);
    assertEquals("username", getScalarValue(usernameTuple.getKeyNode()));
    assertEquals("admin", getScalarValue(usernameTuple.getValueNode()));

    // password: secret # change in production
    NodeTuple passwordTuple = database.getValue().get(3);
    assertEquals("password", getScalarValue(passwordTuple.getKeyNode()));
    assertEquals("secret", getScalarValue(passwordTuple.getValueNode()));
    assertNotNull(passwordTuple.getValueNode().getInLineComments());

    // connection: # connection pool settings
    NodeTuple connectionTuple = database.getValue().get(4);
    assertEquals("connection", getScalarValue(connectionTuple.getKeyNode()));
    // Key inline comment: "# connection pool settings"
    assertNotNull(connectionTuple.getKeyNode().getInLineComments());
    assertTrue(connectionTuple.getValueNode() instanceof MappingNode);

    MappingNode connection = (MappingNode) connectionTuple.getValueNode();
    assertEquals(2, connection.getValue().size()); // pool_size, timeout

    // pool_size: # max connections 10
    NodeTuple poolSizeTuple = connection.getValue().get(0);
    assertEquals("pool_size", getScalarValue(poolSizeTuple.getKeyNode()));
    assertNotNull(poolSizeTuple.getKeyNode().getInLineComments());
    assertEquals("10", getScalarValue(poolSizeTuple.getValueNode()));

    // timeout: # in seconds 30 # 30 seconds (key inline + value + value inline)
    NodeTuple timeoutTuple = connection.getValue().get(1);
    assertEquals("timeout", getScalarValue(timeoutTuple.getKeyNode()));
    assertNotNull(timeoutTuple.getKeyNode().getInLineComments());
    assertEquals("30", getScalarValue(timeoutTuple.getValueNode()));
    assertNotNull(timeoutTuple.getValueNode().getInLineComments());
    assertFalse(timeoutTuple.getValueNode().getInLineComments().isEmpty());

    // replicas: # read replicas
    NodeTuple replicasTuple = database.getValue().get(5);
    assertEquals("replicas", getScalarValue(replicasTuple.getKeyNode()));
    assertNotNull(replicasTuple.getKeyNode().getInLineComments());
    assertTrue(replicasTuple.getValueNode() instanceof SequenceNode);

    SequenceNode replicas = (SequenceNode) replicasTuple.getValueNode();
    assertEquals(2, replicas.getValue().size());

    // First replica: host: replica1.local # first replica
    assertTrue(replicas.getValue().get(0) instanceof MappingNode);
    MappingNode replica1 = (MappingNode) replicas.getValue().get(0);
    assertEquals("host", getScalarValue(replica1.getValue().get(0).getKeyNode()));
    assertEquals("replica1.local", getScalarValue(replica1.getValue().get(0).getValueNode()));
    assertNotNull(replica1.getValue().get(0).getValueNode().getInLineComments());

    // Second replica: host: # second replica address replica2.local
    assertTrue(replicas.getValue().get(1) instanceof MappingNode);
    MappingNode replica2 = (MappingNode) replicas.getValue().get(1);
    assertEquals("host", getScalarValue(replica2.getValue().get(0).getKeyNode()));
    // Key inline comment on host
    assertNotNull(replica2.getValue().get(0).getKeyNode().getInLineComments());
    assertEquals("replica2.local", getScalarValue(replica2.getValue().get(0).getValueNode()));

    // allowed_ips: # whitelist
    NodeTuple allowedIpsTuple = database.getValue().get(6);
    assertEquals("allowed_ips", getScalarValue(allowedIpsTuple.getKeyNode()));
    assertNotNull(allowedIpsTuple.getKeyNode().getInLineComments());
    assertTrue(allowedIpsTuple.getValueNode() instanceof SequenceNode);

    SequenceNode allowedIps = (SequenceNode) allowedIpsTuple.getValueNode();
    assertEquals(3, allowedIps.getValue().size());

    // - 192.168.1.1 # main server
    assertEquals("192.168.1.1", getScalarValue(allowedIps.getValue().get(0)));
    assertNotNull(allowedIps.getValue().get(0).getInLineComments());

    // - 192.168.1.2 # backup server
    assertEquals("192.168.1.2", getScalarValue(allowedIps.getValue().get(1)));
    assertNotNull(allowedIps.getValue().get(1).getInLineComments());

    // # internal network (block comment) - 10.0.0.1
    assertEquals("10.0.0.1", getScalarValue(allowedIps.getValue().get(2)));
    assertNotNull(allowedIps.getValue().get(2).getBlockComments());
    assertFalse(allowedIps.getValue().get(2).getBlockComments().isEmpty());

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

  @Test
  void testValueStartingWithHashRoundTrip() throws IOException {
    // Test that values starting with # are correctly escaped and unescaped
    String yaml = """
        message: '# Look like a comment'
        hashtag: '#trending'
        normal: regular value
        mixed: 'Hello # world'
        """;
    Workbook workbook = YamlWorkbook.toWorkbook(yaml);

    YamlWorkbookReader reader = YamlWorkbookReader.builder().build();
    List<Node> nodes = reader.fromWorkbook(workbook);

    assertEquals(1, nodes.size());
    MappingNode root = (MappingNode) nodes.get(0);
    assertEquals(4, root.getValue().size());

    // Values starting with # should be preserved correctly
    assertEquals("# Look like a comment", getScalarValue(root.getValue().get(0).getValueNode()));
    assertEquals("#trending", getScalarValue(root.getValue().get(1).getValueNode()));
    assertEquals("regular value", getScalarValue(root.getValue().get(2).getValueNode()));
    // # in the middle of the value should NOT be escaped
    assertEquals("Hello # world", getScalarValue(root.getValue().get(3).getValueNode()));

    workbook.close();
  }

  @Test
  void testSequenceItemStartingWithHash() throws IOException {
    // Test sequence items that start with #
    String yaml = """
        items:
          - '# first item'
          - normal item
          - '# another hash item'
        """;
    Workbook workbook = YamlWorkbook.toWorkbook(yaml);

    YamlWorkbookReader reader = YamlWorkbookReader.builder().build();
    List<Node> nodes = reader.fromWorkbook(workbook);

    MappingNode root = (MappingNode) nodes.get(0);
    SequenceNode items = (SequenceNode) root.getValue().get(0).getValueNode();

    assertEquals(3, items.getValue().size());
    assertEquals("# first item", getScalarValue(items.getValue().get(0)));
    assertEquals("normal item", getScalarValue(items.getValue().get(1)));
    assertEquals("# another hash item", getScalarValue(items.getValue().get(2)));

    workbook.close();
  }

  @Test
  void testValueStartingWithBackslash() throws IOException {
    // Test that values starting with \ are correctly double-escaped
    String yaml = """
        path: '\\some\\path'
        escaped_hash: '\\# not a comment'
        """;
    Workbook workbook = YamlWorkbook.toWorkbook(yaml);

    YamlWorkbookReader reader = YamlWorkbookReader.builder().build();
    List<Node> nodes = reader.fromWorkbook(workbook);

    MappingNode root = (MappingNode) nodes.get(0);

    // Backslash at start should be preserved
    assertEquals("\\some\\path", getScalarValue(root.getValue().get(0).getValueNode()));
    assertEquals("\\# not a comment", getScalarValue(root.getValue().get(1).getValueNode()));

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
