package com.github.wnameless.workbook.yamlworkbook;

import static org.junit.jupiter.api.Assertions.*;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.List;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.jupiter.api.BeforeAll;
import org.junit.jupiter.api.Test;
import org.yaml.snakeyaml.nodes.MappingNode;
import org.yaml.snakeyaml.nodes.Node;
import org.yaml.snakeyaml.nodes.NodeTuple;
import org.yaml.snakeyaml.nodes.ScalarNode;
import org.yaml.snakeyaml.nodes.SequenceNode;

class YamlWorkbookReaderTest {

  private static final boolean WRITE_EXCEL_FILES = true;
  private static final Path OUTPUT_DIR = Paths.get("target/test-excel");

  @BeforeAll
  static void setUpOnce() throws IOException {
    if (WRITE_EXCEL_FILES) {
      Files.createDirectories(OUTPUT_DIR);
    }
  }

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
    assertEquals(7, database.getValue().size()); // host, port, username, password, connection,
                                                 // replicas, allowed_ips

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

  private void writeExcelFile(Workbook workbook, String filename) throws IOException {
    if (WRITE_EXCEL_FILES) {
      Path filePath = OUTPUT_DIR.resolve(filename);
      try (FileOutputStream fos = new FileOutputStream(filePath.toFile())) {
        workbook.write(fos);
      }
    }
  }

  private String loadJsonSchema(String resourcePath) throws IOException {
    try (InputStream is = getClass().getClassLoader().getResourceAsStream(resourcePath)) {
      if (is == null) {
        throw new IOException("Resource not found: " + resourcePath);
      }
      return new String(is.readAllBytes(), StandardCharsets.UTF_8);
    }
  }

  // ==================== DATA_COLLECT Mode Roundtrip Tests ====================

  @Test
  void testDataCollectModeEnumDropdownRoundtrip() throws IOException {
    // Load the JSON schema with small and large enums
    String jsonSchema = loadJsonSchema("schema/enum-dropdown-test.json");

    // Write with DATA_COLLECT mode and hidden sheets for large enums
    // Note: DATA_COLLECT mode generates workbook from schema, not from YAML input
    DataCollectConfig config = DataCollectConfig.builder()
        .useHiddenSheetsForLongEnums(true)
        .build();

    YamlWorkbookWriter writer = YamlWorkbookWriter.builder()
        .printMode(PrintMode.DATA_COLLECT)
        .dataCollectConfig(config)
        .jsonSchema(jsonSchema)
        .build();

    // Use toWorkbook() without parameters for DATA_COLLECT mode
    Workbook workbook = writer.toWorkbook();
    assertNotNull(workbook);

    // Verify we have both visible and hidden sheets
    // Sheet1 (visible) + Sheet1Hidden (hidden for large enum)
    assertTrue(workbook.getNumberOfSheets() >= 1);

    // Find the hidden sheet (country enum exceeds 256 chars)
    boolean hasHiddenSheet = false;
    for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
      if (workbook.isSheetHidden(i)) {
        hasHiddenSheet = true;
        break;
      }
    }
    assertTrue(hasHiddenSheet, "Expected a hidden sheet for large enum dropdown");

    // Write file for inspection
    writeExcelFile(workbook, "data_collect_enum_roundtrip.xlsx");

    // Simulate user filling in data by selecting from dropdowns
    // Find rows by their key cell values (property order may vary)
    Sheet sheet = workbook.getSheetAt(0);
    for (int i = 1; i <= sheet.getLastRowNum(); i++) {
      Row row = sheet.getRow(i);
      if (row == null) continue;
      Cell keyCell = row.getCell(0);
      if (keyCell == null) continue;
      String key = keyCell.getStringCellValue();
      Cell valueCell = row.getCell(1);
      if (valueCell == null) continue;

      // Set enum display values (simulating dropdown selection)
      switch (key) {
        case "Status" -> valueCell.setCellValue("Active");
        case "Priority Level" -> valueCell.setCellValue("High");
        case "Country" -> valueCell.setCellValue("United States");
        case "Department" -> valueCell.setCellValue("Engineering");
      }
    }

    // Read with DATA_COLLECT mode (should recover original enum values)
    YamlWorkbookReader reader = YamlWorkbookReader.builder()
        .printMode(PrintMode.DATA_COLLECT)
        .build();
    List<Node> nodes = reader.fromWorkbook(workbook);

    assertEquals(1, nodes.size());
    MappingNode root = (MappingNode) nodes.get(0);
    assertEquals(4, root.getValue().size());

    // Build a map of recovered key-value pairs for flexible verification
    java.util.Map<String, String> recovered = new java.util.HashMap<>();
    for (NodeTuple tuple : root.getValue()) {
      String key = getScalarValue(tuple.getKeyNode());
      String value = getScalarValue(tuple.getValueNode());
      recovered.put(key, value);
    }

    // Verify original enum values are recovered (not enumNames)
    assertEquals("active", recovered.get("status"), "status should map to 'active'");
    assertEquals("2", recovered.get("priority"), "priority should map to '2'");
    assertEquals("US", recovered.get("country"), "country should map to 'US'");
    assertEquals("eng", recovered.get("department"), "department should map to 'eng'");

    workbook.close();
  }

  @Test
  void testDataCollectModeEnumSelectionRoundtrip() throws IOException {
    // Test simulating user selection: write workbook, modify cell values, read back
    String jsonSchema = loadJsonSchema("schema/enum-dropdown-test.json");

    DataCollectConfig config = DataCollectConfig.builder()
        .useHiddenSheetsForLongEnums(true)
        .build();

    YamlWorkbookWriter writer = YamlWorkbookWriter.builder()
        .printMode(PrintMode.DATA_COLLECT)
        .dataCollectConfig(config)
        .jsonSchema(jsonSchema)
        .build();

    // Use toWorkbook() without parameters for DATA_COLLECT mode
    Workbook workbook = writer.toWorkbook();

    // Simulate user selecting different values from dropdowns
    // Find rows by their key cell values (property order may vary)
    Sheet sheet = workbook.getSheetAt(0);
    for (int i = 1; i <= sheet.getLastRowNum(); i++) {
      Row row = sheet.getRow(i);
      if (row == null) continue;
      Cell keyCell = row.getCell(0);
      if (keyCell == null) continue;
      String key = keyCell.getStringCellValue();
      Cell valueCell = row.getCell(1);
      if (valueCell == null) continue;

      // Set different enum display values (simulating dropdown selection)
      switch (key) {
        case "Status" -> valueCell.setCellValue("Pending");
        case "Priority Level" -> valueCell.setCellValue("Low");
        case "Country" -> valueCell.setCellValue("Japan");
        case "Department" -> valueCell.setCellValue("Finance");
      }
    }

    // Write modified file for inspection
    writeExcelFile(workbook, "data_collect_enum_selection_roundtrip.xlsx");

    // Read back with DATA_COLLECT mode
    YamlWorkbookReader reader = YamlWorkbookReader.builder()
        .printMode(PrintMode.DATA_COLLECT)
        .build();
    List<Node> nodes = reader.fromWorkbook(workbook);

    assertEquals(1, nodes.size());
    MappingNode root = (MappingNode) nodes.get(0);
    assertEquals(4, root.getValue().size());

    // Build a map of recovered key-value pairs for flexible verification
    java.util.Map<String, String> recovered = new java.util.HashMap<>();
    for (NodeTuple tuple : root.getValue()) {
      String key = getScalarValue(tuple.getKeyNode());
      String value = getScalarValue(tuple.getValueNode());
      recovered.put(key, value);
    }

    // Verify original enum values are recovered from user selections
    assertEquals("pending", recovered.get("status"), "status: Pending -> pending");
    assertEquals("4", recovered.get("priority"), "priority: Low -> 4");
    assertEquals("JP", recovered.get("country"), "country: Japan -> JP (via hidden sheet)");
    assertEquals("fin", recovered.get("department"), "department: Finance -> fin");

    workbook.close();
  }

  @Test
  void testDataCollectModeSmallEnumWithoutHiddenSheet() throws IOException {
    // Test that small enums don't create hidden sheets when useHiddenSheetsForLongEnums is false
    // Use a schema with only small enums
    String jsonSchema = """
        {
          "$schema": "http://json-schema.org/draft-07/schema#",
          "type": "object",
          "properties": {
            "status": {
              "type": "string",
              "title": "Status",
              "enum": ["active", "inactive", "pending"],
              "enumNames": ["Active", "Inactive", "Pending"]
            },
            "priority": {
              "type": "integer",
              "title": "Priority Level",
              "enum": [1, 2, 3],
              "enumNames": ["High", "Medium", "Low"]
            }
          }
        }
        """;

    // Small enums should not create hidden sheets
    DataCollectConfig config = DataCollectConfig.builder()
        .useHiddenSheetsForLongEnums(true)  // Even if true, small enums don't need hidden sheets
        .build();

    YamlWorkbookWriter writer = YamlWorkbookWriter.builder()
        .printMode(PrintMode.DATA_COLLECT)
        .dataCollectConfig(config)
        .jsonSchema(jsonSchema)
        .build();

    Workbook workbook = writer.toWorkbook();

    // Verify no hidden sheets (small enums use explicit list constraint)
    boolean hasHiddenSheet = false;
    for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
      if (workbook.isSheetHidden(i)) {
        hasHiddenSheet = true;
        break;
      }
    }
    assertFalse(hasHiddenSheet, "Small enums should not create hidden sheets");

    writeExcelFile(workbook, "data_collect_small_enum_no_hidden.xlsx");

    // Simulate user selection (find by key name)
    Sheet sheet = workbook.getSheetAt(0);
    for (int i = 1; i <= sheet.getLastRowNum(); i++) {
      Row row = sheet.getRow(i);
      if (row == null) continue;
      Cell keyCell = row.getCell(0);
      if (keyCell == null) continue;
      String key = keyCell.getStringCellValue();
      Cell valueCell = row.getCell(1);
      if (valueCell == null) continue;

      switch (key) {
        case "Status" -> valueCell.setCellValue("Active");
        case "Priority Level" -> valueCell.setCellValue("Medium");
      }
    }

    // Read back and verify
    YamlWorkbookReader reader = YamlWorkbookReader.builder()
        .printMode(PrintMode.DATA_COLLECT)
        .build();
    List<Node> nodes = reader.fromWorkbook(workbook);

    assertEquals(1, nodes.size());
    MappingNode root = (MappingNode) nodes.get(0);
    assertEquals(2, root.getValue().size());

    // Build a map of recovered key-value pairs
    java.util.Map<String, String> recovered = new java.util.HashMap<>();
    for (NodeTuple tuple : root.getValue()) {
      String key = getScalarValue(tuple.getKeyNode());
      String value = getScalarValue(tuple.getValueNode());
      recovered.put(key, value);
    }

    // Verify values are recovered
    assertEquals("active", recovered.get("status"));
    assertEquals("2", recovered.get("priority"));

    workbook.close();
  }

  @Test
  void testDataCollectModeLargeEnumTruncation() throws IOException {
    // Test large enum truncation when hidden sheets are disabled
    String jsonSchema = loadJsonSchema("schema/enum-dropdown-test.json");

    // Disable hidden sheets - large enum will be truncated
    DataCollectConfig config = DataCollectConfig.builder()
        .useHiddenSheetsForLongEnums(false)
        .build();

    YamlWorkbookWriter writer = YamlWorkbookWriter.builder()
        .printMode(PrintMode.DATA_COLLECT)
        .dataCollectConfig(config)
        .jsonSchema(jsonSchema)
        .build();

    Workbook workbook = writer.toWorkbook();

    // Verify no hidden sheets
    boolean hasHiddenSheet = false;
    for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
      if (workbook.isSheetHidden(i)) {
        hasHiddenSheet = true;
        break;
      }
    }
    assertFalse(hasHiddenSheet, "Hidden sheets should be disabled");

    writeExcelFile(workbook, "data_collect_large_enum_truncated.xlsx");

    // Simulate user selections (find by key name)
    Sheet sheet = workbook.getSheetAt(0);
    for (int i = 1; i <= sheet.getLastRowNum(); i++) {
      Row row = sheet.getRow(i);
      if (row == null) continue;
      Cell keyCell = row.getCell(0);
      if (keyCell == null) continue;
      String key = keyCell.getStringCellValue();
      Cell valueCell = row.getCell(1);
      if (valueCell == null) continue;

      switch (key) {
        case "Status" -> valueCell.setCellValue("Active");
        case "Priority Level" -> valueCell.setCellValue("Critical");
        case "Country" -> valueCell.setCellValue("United States");
        case "Department" -> valueCell.setCellValue("Engineering");
      }
    }

    // Read back - the dropdown is truncated but value should still be recoverable
    // if it's in the truncated list
    YamlWorkbookReader reader = YamlWorkbookReader.builder()
        .printMode(PrintMode.DATA_COLLECT)
        .build();
    List<Node> nodes = reader.fromWorkbook(workbook);

    assertEquals(1, nodes.size());
    MappingNode root = (MappingNode) nodes.get(0);

    // Build a map of recovered key-value pairs
    java.util.Map<String, String> recovered = new java.util.HashMap<>();
    for (NodeTuple tuple : root.getValue()) {
      String key = getScalarValue(tuple.getKeyNode());
      String value = getScalarValue(tuple.getValueNode());
      recovered.put(key, value);
    }

    // US should be recoverable since "United States" is in first few options of truncated list
    assertEquals("US", recovered.get("country"), "country: United States -> US");

    workbook.close();
  }

  @Test
  void testDataCollectModeKeyTitleReplacement() throws IOException {
    // Test that key titles from JSON schema are used and original keys recovered
    String jsonSchema = loadJsonSchema("schema/enum-dropdown-test.json");

    DataCollectConfig config = DataCollectConfig.builder()
        .useHiddenSheetsForLongEnums(true)
        .build();

    YamlWorkbookWriter writer = YamlWorkbookWriter.builder()
        .printMode(PrintMode.DATA_COLLECT)
        .dataCollectConfig(config)
        .jsonSchema(jsonSchema)
        .build();

    Workbook workbook = writer.toWorkbook();

    // Verify the key cells show titles (not original keys) by checking all rows
    Sheet sheet = workbook.getSheetAt(0);
    java.util.Set<String> foundTitles = new java.util.HashSet<>();
    for (int i = 1; i <= sheet.getLastRowNum(); i++) {
      Row row = sheet.getRow(i);
      if (row == null) continue;
      Cell keyCell = row.getCell(0);
      if (keyCell == null) continue;
      foundTitles.add(keyCell.getStringCellValue());
    }
    assertTrue(foundTitles.contains("Status"), "Should have 'Status' title");
    assertTrue(foundTitles.contains("Priority Level"), "Should have 'Priority Level' title");
    assertTrue(foundTitles.contains("Country"), "Should have 'Country' title");
    assertTrue(foundTitles.contains("Department"), "Should have 'Department' title");

    writeExcelFile(workbook, "data_collect_key_title.xlsx");

    // Simulate user selections (find by key name)
    for (int i = 1; i <= sheet.getLastRowNum(); i++) {
      Row row = sheet.getRow(i);
      if (row == null) continue;
      Cell keyCell = row.getCell(0);
      if (keyCell == null) continue;
      String key = keyCell.getStringCellValue();
      Cell valueCell = row.getCell(1);
      if (valueCell == null) continue;

      switch (key) {
        case "Status" -> valueCell.setCellValue("Active");
        case "Priority Level" -> valueCell.setCellValue("Medium");
        case "Country" -> valueCell.setCellValue("United States");
        case "Department" -> valueCell.setCellValue("Engineering");
      }
    }

    // Read back and verify original keys are recovered
    YamlWorkbookReader reader = YamlWorkbookReader.builder()
        .printMode(PrintMode.DATA_COLLECT)
        .build();
    List<Node> nodes = reader.fromWorkbook(workbook);

    MappingNode root = (MappingNode) nodes.get(0);

    // Build a set of recovered keys
    java.util.Set<String> recoveredKeys = new java.util.HashSet<>();
    for (NodeTuple tuple : root.getValue()) {
      recoveredKeys.add(getScalarValue(tuple.getKeyNode()));
    }

    // Original keys should be recovered from cell comments
    assertTrue(recoveredKeys.contains("status"), "Should recover 'status' key");
    assertTrue(recoveredKeys.contains("priority"), "Should recover 'priority' key");
    assertTrue(recoveredKeys.contains("country"), "Should recover 'country' key");
    assertTrue(recoveredKeys.contains("department"), "Should recover 'department' key");

    workbook.close();
  }

  // ==================== WORKBOOK_READABLE Mode Roundtrip Tests ====================

  @Test
  void testReadableModeValueReplacementRoundtrip() throws IOException {
    // Test that value replaced by inline comment is recovered via cell comment
    String yaml = """
        name: John Doe
        age: 30 # Age in Years
        city: New York
        """;

    // Write with WORKBOOK_READABLE mode (value replaced by comment)
    YamlWorkbookWriter writer =
        YamlWorkbookWriter.builder().printMode(PrintMode.WORKBOOK_READABLE).build();
    Workbook workbook = writer.toWorkbook(new java.io.StringReader(yaml));

    // Read with WORKBOOK_READABLE mode (should recover original value from cell comment)
    YamlWorkbookReader reader =
        YamlWorkbookReader.builder().printMode(PrintMode.WORKBOOK_READABLE).build();
    List<Node> nodes = reader.fromWorkbook(workbook);

    assertEquals(1, nodes.size());
    MappingNode root = (MappingNode) nodes.get(0);
    assertEquals(3, root.getValue().size());

    // name: John Doe (no comment, unchanged)
    assertEquals("name", getScalarValue(root.getValue().get(0).getKeyNode()));
    assertEquals("John Doe", getScalarValue(root.getValue().get(0).getValueNode()));

    // age: 30 (original value recovered from cell comment, not "Age in Years")
    assertEquals("age", getScalarValue(root.getValue().get(1).getKeyNode()));
    assertEquals("30", getScalarValue(root.getValue().get(1).getValueNode()));

    // city: New York (no comment, unchanged)
    assertEquals("city", getScalarValue(root.getValue().get(2).getKeyNode()));
    assertEquals("New York", getScalarValue(root.getValue().get(2).getValueNode()));

    workbook.close();
  }

  @Test
  void testReadableModeKeyReplacementRoundtrip() throws IOException {
    // Test that key replaced by inline comment is recovered via cell comment
    String yaml = """
        user_name: # Display Name
          John Doe
        user_age: 30
        """;

    // Write with WORKBOOK_READABLE mode and keyComment=DISPLAY_NAME (default)
    YamlWorkbookWriter writer =
        YamlWorkbookWriter.builder().printMode(PrintMode.WORKBOOK_READABLE).build();
    Workbook workbook = writer.toWorkbook(new java.io.StringReader(yaml));

    // Read with WORKBOOK_READABLE mode (should recover original key from cell comment)
    YamlWorkbookReader reader =
        YamlWorkbookReader.builder().printMode(PrintMode.WORKBOOK_READABLE).build();
    List<Node> nodes = reader.fromWorkbook(workbook);

    assertEquals(1, nodes.size());
    MappingNode root = (MappingNode) nodes.get(0);
    assertEquals(2, root.getValue().size());

    // user_name: John Doe (original key recovered from cell comment, not "Display Name")
    assertEquals("user_name", getScalarValue(root.getValue().get(0).getKeyNode()));
    assertEquals("John Doe", getScalarValue(root.getValue().get(0).getValueNode()));

    // user_age: 30 (no comment, unchanged)
    assertEquals("user_age", getScalarValue(root.getValue().get(1).getKeyNode()));
    assertEquals("30", getScalarValue(root.getValue().get(1).getValueNode()));

    workbook.close();
  }

  @Test
  void testReadableModeBlockCommentRoundtrip() throws IOException {
    // Test that block comment used as DISPLAY_NAME header is recovered
    String yaml = """
        user:
          # User Settings
          settings:
            theme: dark
        """;

    // Write with WORKBOOK_READABLE mode and objectComment=DISPLAY_NAME (default)
    DisplayModeConfig config =
        DisplayModeConfig.builder().keyValuePairComment(CommentVisibility.HIDDEN).build();

    YamlWorkbookWriter writer = YamlWorkbookWriter.builder().printMode(PrintMode.WORKBOOK_READABLE)
        .displayModeConfig(config).build();
    Workbook workbook = writer.toWorkbook(new java.io.StringReader(yaml));

    // Read with WORKBOOK_READABLE mode
    YamlWorkbookReader reader =
        YamlWorkbookReader.builder().printMode(PrintMode.WORKBOOK_READABLE).build();
    List<Node> nodes = reader.fromWorkbook(workbook);

    assertEquals(1, nodes.size());
    MappingNode root = (MappingNode) nodes.get(0);

    // user:
    NodeTuple userTuple = root.getValue().get(0);
    assertEquals("user", getScalarValue(userTuple.getKeyNode()));
    assertTrue(userTuple.getValueNode() instanceof MappingNode);

    // MappingNode user = (MappingNode) userTuple.getValueNode();
    // The block comment "# User Settings" should be preserved as a comment
    // (cell comment contains "# User Settings" which is recognized as a comment)

    workbook.close();
  }

  @Test
  void testReadableModeMixedRoundtrip() throws IOException {
    // Test complex scenario with both key and value replacements
    String yaml = """
        db_host: # Server Address
          localhost # Primary Host
        db_port: 5432 # Port Number
        """;

    YamlWorkbookWriter writer =
        YamlWorkbookWriter.builder().printMode(PrintMode.WORKBOOK_READABLE).build();
    Workbook workbook = writer.toWorkbook(new java.io.StringReader(yaml));

    YamlWorkbookReader reader =
        YamlWorkbookReader.builder().printMode(PrintMode.WORKBOOK_READABLE).build();
    List<Node> nodes = reader.fromWorkbook(workbook);

    assertEquals(1, nodes.size());
    MappingNode root = (MappingNode) nodes.get(0);
    assertEquals(2, root.getValue().size());

    // db_host: localhost (both key and value recovered from cell comments)
    assertEquals("db_host", getScalarValue(root.getValue().get(0).getKeyNode()));
    assertEquals("localhost", getScalarValue(root.getValue().get(0).getValueNode()));

    // db_port: 5432 (value recovered from cell comment)
    assertEquals("db_port", getScalarValue(root.getValue().get(1).getKeyNode()));
    assertEquals("5432", getScalarValue(root.getValue().get(1).getValueNode()));

    workbook.close();
  }

  @Test
  void testReadableModeNoCommentUnchanged() throws IOException {
    // Test that cells without comments work normally in WORKBOOK_READABLE mode
    String yaml = """
        name: John Doe
        age: 30
        city: New York
        """;

    YamlWorkbookWriter writer =
        YamlWorkbookWriter.builder().printMode(PrintMode.WORKBOOK_READABLE).build();
    Workbook workbook = writer.toWorkbook(new java.io.StringReader(yaml));

    YamlWorkbookReader reader =
        YamlWorkbookReader.builder().printMode(PrintMode.WORKBOOK_READABLE).build();
    List<Node> nodes = reader.fromWorkbook(workbook);

    assertEquals(1, nodes.size());
    MappingNode root = (MappingNode) nodes.get(0);
    assertEquals(3, root.getValue().size());

    // All values should be unchanged (no inline comments in YAML)
    assertEquals("name", getScalarValue(root.getValue().get(0).getKeyNode()));
    assertEquals("John Doe", getScalarValue(root.getValue().get(0).getValueNode()));

    assertEquals("age", getScalarValue(root.getValue().get(1).getKeyNode()));
    assertEquals("30", getScalarValue(root.getValue().get(1).getValueNode()));

    assertEquals("city", getScalarValue(root.getValue().get(2).getKeyNode()));
    assertEquals("New York", getScalarValue(root.getValue().get(2).getValueNode()));

    workbook.close();
  }

  @Test
  void testYamlOrientedModeIgnoresCellComments() throws IOException {
    // Test that YAML_ORIENTED mode ignores cell comments (reads displayed value)
    String yaml = """
        age: 30 # Age in Years
        """;

    // Write with WORKBOOK_READABLE mode (creates cell comments)
    YamlWorkbookWriter writer =
        YamlWorkbookWriter.builder().printMode(PrintMode.WORKBOOK_READABLE).build();
    Workbook workbook = writer.toWorkbook(new java.io.StringReader(yaml));

    // Read with YAML_ORIENTED mode (should ignore cell comments, read displayed value)
    YamlWorkbookReader reader =
        YamlWorkbookReader.builder().printMode(PrintMode.YAML_ORIENTED).build();
    List<Node> nodes = reader.fromWorkbook(workbook);

    assertEquals(1, nodes.size());
    MappingNode root = (MappingNode) nodes.get(0);

    // In YAML_ORIENTED mode, reader should see "Age in Years" (the displayed value)
    // not "30" (the original value in cell comment)
    assertEquals("age", getScalarValue(root.getValue().get(0).getKeyNode()));
    assertEquals("Age in Years", getScalarValue(root.getValue().get(0).getValueNode()));

    workbook.close();
  }

}
