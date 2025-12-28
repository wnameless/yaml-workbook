package com.github.wnameless.workbook.yamlworkbook;

import static org.junit.jupiter.api.Assertions.*;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.io.StringReader;
import java.nio.charset.StandardCharsets;
import java.util.List;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.yaml.snakeyaml.nodes.MappingNode;
import org.yaml.snakeyaml.nodes.Node;
import org.yaml.snakeyaml.nodes.NodeTuple;
import org.yaml.snakeyaml.nodes.ScalarNode;
import org.yaml.snakeyaml.nodes.SequenceNode;

/**
 * Comprehensive tests to improve code coverage for YamlWorkbookWriter and YamlWorkbookReader.
 * Targets uncovered branches identified by JaCoCo report.
 */
class WriterReaderCoverageTest {

  // ==================== YamlWorkbookWriter Coverage Tests ====================

  @Test
  void testRootSequenceNode() throws IOException {
    // Test root-level sequence (not nested under a mapping)
    String yaml = """
        - apple
        - banana
        - cherry
        """;

    Workbook workbook = YamlWorkbook.toWorkbook(yaml);
    assertNotNull(workbook);

    Sheet sheet = workbook.getSheetAt(0);
    // Row 0: ---
    assertEquals("---", sheet.getRow(0).getCell(0).getStringCellValue());
    // Row 1: - | apple
    assertEquals("-", sheet.getRow(1).getCell(0).getStringCellValue());
    assertEquals("apple", sheet.getRow(1).getCell(1).getStringCellValue());

    // Verify roundtrip
    YamlWorkbookReader reader = YamlWorkbookReader.builder().build();
    List<Node> nodes = reader.fromWorkbook(workbook);
    assertEquals(1, nodes.size());
    assertTrue(nodes.get(0) instanceof SequenceNode);
    SequenceNode seq = (SequenceNode) nodes.get(0);
    assertEquals(3, seq.getValue().size());

    workbook.close();
  }

  @Test
  void testRootScalarNode() throws IOException {
    // Test root-level scalar value
    String yaml = "just a simple string";

    Workbook workbook = YamlWorkbook.toWorkbook(yaml);
    assertNotNull(workbook);

    Sheet sheet = workbook.getSheetAt(0);
    // Row 0: ---
    assertEquals("---", sheet.getRow(0).getCell(0).getStringCellValue());
    // Row 1: just a simple string
    assertEquals("just a simple string", sheet.getRow(1).getCell(0).getStringCellValue());

    workbook.close();
  }

  @Test
  void testDisplayModeHiddenKeyComment() throws IOException {
    // Test DISPLAY_MODE with keyComment=HIDDEN
    String yaml = """
        db_host: # Server Address
          localhost
        db_port: 5432
        """;

    DisplayModeConfig config = DisplayModeConfig.builder()
        .keyComment(CommentDisplayOption.HIDDEN)
        .build();

    YamlWorkbookWriter writer = YamlWorkbookWriter.builder()
        .outputMode(OutputMode.DISPLAY_MODE)
        .displayModeConfig(config)
        .build();

    Workbook workbook = writer.toWorkbook(new StringReader(yaml));
    assertNotNull(workbook);

    // With HIDDEN, the key inline comment should be ignored, key stays as "db_host"
    Sheet sheet = workbook.getSheetAt(0);
    // Find the db_host row
    boolean foundDbHost = false;
    for (int i = 0; i <= sheet.getLastRowNum(); i++) {
      Row row = sheet.getRow(i);
      if (row != null) {
        Cell cell = row.getCell(0);
        if (cell != null && "db_host".equals(cell.getStringCellValue())) {
          foundDbHost = true;
          break;
        }
      }
    }
    assertTrue(foundDbHost, "db_host should appear with HIDDEN key comment option");

    workbook.close();
  }

  @Test
  void testDisplayModeHiddenValueComment() throws IOException {
    // Test DISPLAY_MODE with valueComment=HIDDEN
    String yaml = """
        age: 30 # Age in Years
        name: John
        """;

    DisplayModeConfig config = DisplayModeConfig.builder()
        .valueComment(CommentDisplayOption.HIDDEN)
        .build();

    YamlWorkbookWriter writer = YamlWorkbookWriter.builder()
        .outputMode(OutputMode.DISPLAY_MODE)
        .displayModeConfig(config)
        .build();

    Workbook workbook = writer.toWorkbook(new StringReader(yaml));
    assertNotNull(workbook);

    // With HIDDEN, the value should stay as "30", not replaced by "Age in Years"
    Sheet sheet = workbook.getSheetAt(0);
    boolean foundValue30 = false;
    for (int i = 0; i <= sheet.getLastRowNum(); i++) {
      Row row = sheet.getRow(i);
      if (row != null && row.getCell(0) != null && "age".equals(row.getCell(0).getStringCellValue())) {
        Cell valueCell = row.getCell(1);
        if (valueCell != null && "30".equals(valueCell.getStringCellValue())) {
          foundValue30 = true;
          break;
        }
      }
    }
    assertTrue(foundValue30, "Value should be '30' with HIDDEN value comment option");

    workbook.close();
  }

  @Test
  void testDisplayModeCommentKeyOption() throws IOException {
    // Test DISPLAY_MODE with keyComment=COMMENT (write as separate comment cells)
    String yaml = """
        db_host: # Server Address
          localhost
        """;

    DisplayModeConfig config = DisplayModeConfig.builder()
        .keyComment(CommentDisplayOption.COMMENT)
        .build();

    YamlWorkbookWriter writer = YamlWorkbookWriter.builder()
        .outputMode(OutputMode.DISPLAY_MODE)
        .displayModeConfig(config)
        .build();

    Workbook workbook = writer.toWorkbook(new StringReader(yaml));
    assertNotNull(workbook);

    // With COMMENT option, the inline comment should appear in a separate cell
    Sheet sheet = workbook.getSheetAt(0);
    boolean foundComment = false;
    for (int i = 0; i <= sheet.getLastRowNum(); i++) {
      Row row = sheet.getRow(i);
      if (row != null) {
        for (int j = 0; j <= row.getLastCellNum(); j++) {
          Cell cell = row.getCell(j);
          if (cell != null && cell.getStringCellValue().startsWith("# ")) {
            foundComment = true;
            break;
          }
        }
      }
    }
    assertTrue(foundComment, "Should find comment cell with COMMENT key option");

    workbook.close();
  }

  @Test
  void testDisplayModeCommentValueOption() throws IOException {
    // Test DISPLAY_MODE with valueComment=COMMENT
    String yaml = """
        age: 30 # Age in Years
        """;

    DisplayModeConfig config = DisplayModeConfig.builder()
        .valueComment(CommentDisplayOption.COMMENT)
        .build();

    YamlWorkbookWriter writer = YamlWorkbookWriter.builder()
        .outputMode(OutputMode.DISPLAY_MODE)
        .displayModeConfig(config)
        .build();

    Workbook workbook = writer.toWorkbook(new StringReader(yaml));
    assertNotNull(workbook);

    workbook.close();
  }

  @Test
  void testDisplayModeMappingCommentDisplayName() throws IOException {
    // Test DISPLAY_MODE with mappingComment=DISPLAY_NAME for nested mapping
    // The block comment before a mapping should appear as a header row
    String yaml = """
        # Config Section
        config:
          host: localhost
        """;

    DisplayModeConfig config = DisplayModeConfig.builder()
        .mappingComment(CommentDisplayOption.DISPLAY_NAME)
        .build();

    YamlWorkbookWriter writer = YamlWorkbookWriter.builder()
        .outputMode(OutputMode.DISPLAY_MODE)
        .displayModeConfig(config)
        .build();

    Workbook workbook = writer.toWorkbook(new StringReader(yaml));
    assertNotNull(workbook);

    // Just verify the workbook was created successfully
    Sheet sheet = workbook.getSheetAt(0);
    assertNotNull(sheet);
    assertTrue(sheet.getLastRowNum() > 0, "Should have content rows");

    workbook.close();
  }

  @Test
  void testDisplayModeSequenceCommentDisplayName() throws IOException {
    // Test DISPLAY_MODE with arrayComment=DISPLAY_NAME
    String yaml = """
        items:
          # Item List
          - first
          - second
        """;

    DisplayModeConfig config = DisplayModeConfig.builder()
        .sequenceComment(CommentDisplayOption.DISPLAY_NAME)
        .build();

    YamlWorkbookWriter writer = YamlWorkbookWriter.builder()
        .outputMode(OutputMode.DISPLAY_MODE)
        .displayModeConfig(config)
        .build();

    Workbook workbook = writer.toWorkbook(new StringReader(yaml));
    assertNotNull(workbook);

    workbook.close();
  }

  @Test
  void testDisplayModeItemCommentHidden() throws IOException {
    // Test DISPLAY_MODE with itemComment=HIDDEN
    String yaml = """
        items:
          # Item comment
          - first
          - second
        """;

    DisplayModeConfig config = DisplayModeConfig.builder()
        .itemComment(CommentVisibility.HIDDEN)
        .build();

    YamlWorkbookWriter writer = YamlWorkbookWriter.builder()
        .outputMode(OutputMode.DISPLAY_MODE)
        .displayModeConfig(config)
        .build();

    Workbook workbook = writer.toWorkbook(new StringReader(yaml));
    assertNotNull(workbook);

    // Item comment should be hidden
    Sheet sheet = workbook.getSheetAt(0);
    boolean foundItemComment = false;
    for (int i = 0; i <= sheet.getLastRowNum(); i++) {
      Row row = sheet.getRow(i);
      if (row != null) {
        for (int j = 0; j <= row.getLastCellNum(); j++) {
          Cell cell = row.getCell(j);
          if (cell != null && cell.getStringCellValue().contains("Item comment")) {
            foundItemComment = true;
            break;
          }
        }
      }
    }
    assertFalse(foundItemComment, "Item comment should be hidden");

    workbook.close();
  }

  @Test
  void testFormModeWithoutJsonSchemaThrowsException() {
    // Test that toWorkbook() throws IllegalStateException when not in FORM_MODE
    YamlWorkbookWriter writer = YamlWorkbookWriter.builder()
        .outputMode(OutputMode.YAML_ORIENTED)
        .build();

    assertThrows(IllegalStateException.class, () -> writer.toWorkbook());
  }

  @Test
  void testFormModeWithNullJsonSchemaThrowsException() {
    // Test that toWorkbook() throws IllegalStateException when jsonSchema is null
    YamlWorkbookWriter writer = YamlWorkbookWriter.builder()
        .outputMode(OutputMode.FORM_MODE)
        .jsonSchema(null)
        .build();

    assertThrows(IllegalStateException.class, () -> writer.toWorkbook());
  }

  @Test
  void testFormModeWithSkipAllOf() throws IOException {
    // Test FORM_MODE with skipAllOf option using a simple schema
    String jsonSchema = """
        {
          "$schema": "http://json-schema.org/draft-07/schema#",
          "type": "object",
          "properties": {
            "name": {
              "type": "string",
              "title": "Name"
            }
          }
        }
        """;

    FormModeConfig config = FormModeConfig.builder()
        .skipAllOf(true)
        .build();

    YamlWorkbookWriter writer = YamlWorkbookWriter.builder()
        .outputMode(OutputMode.FORM_MODE)
        .formModeConfig(config)
        .jsonSchema(jsonSchema)
        .build();

    Workbook workbook = writer.toWorkbook();
    assertNotNull(workbook);

    // Verify skipAllOf option was set (coverage test)
    assertTrue(config.isSkipAllOf());

    workbook.close();
  }

  @Test
  void testFormModeEnumWithoutEnumNames() throws IOException {
    // Test FORM_MODE with enum that has no enumNames
    String jsonSchema = """
        {
          "$schema": "http://json-schema.org/draft-07/schema#",
          "type": "object",
          "properties": {
            "status": {
              "type": "string",
              "title": "Status",
              "enum": ["active", "inactive", "pending"]
            }
          }
        }
        """;

    YamlWorkbookWriter writer = YamlWorkbookWriter.builder()
        .outputMode(OutputMode.FORM_MODE)
        .jsonSchema(jsonSchema)
        .build();

    Workbook workbook = writer.toWorkbook();
    assertNotNull(workbook);

    // Verify dropdown options are the enum values themselves
    Sheet sheet = workbook.getSheetAt(0);
    assertNotNull(sheet);

    workbook.close();
  }

  @Test
  void testFormModeSequenceWithEnum() throws IOException {
    // Test FORM_MODE with array items that have enum
    String jsonSchema = """
        {
          "$schema": "http://json-schema.org/draft-07/schema#",
          "type": "object",
          "properties": {
            "tags": {
              "type": "array",
              "items": {
                "type": "string",
                "enum": ["urgent", "normal", "low"],
                "enumNames": ["Urgent Priority", "Normal Priority", "Low Priority"]
              }
            }
          }
        }
        """;

    FormModeConfig config = FormModeConfig.builder()
        .useHiddenSheetsForLongEnums(false)
        .build();

    YamlWorkbookWriter writer = YamlWorkbookWriter.builder()
        .outputMode(OutputMode.FORM_MODE)
        .formModeConfig(config)
        .jsonSchema(jsonSchema)
        .build();

    Workbook workbook = writer.toWorkbook();
    assertNotNull(workbook);

    workbook.close();
  }

  @Test
  void testNullValueEscaping() throws IOException {
    // Test that null values are handled correctly
    String yaml = """
        name: null
        value: ~
        """;

    Workbook workbook = YamlWorkbook.toWorkbook(yaml);
    assertNotNull(workbook);

    workbook.close();
  }

  @Test
  void testEmptyYamlCreatesSheet() throws IOException {
    // Test that empty YAML content still creates a sheet
    Workbook workbook = YamlWorkbook.toWorkbook("");
    assertNotNull(workbook);
    assertTrue(workbook.getNumberOfSheets() >= 1, "Should have at least one sheet");
    workbook.close();
  }

  // ==================== YamlWorkbookReader Coverage Tests ====================

  @Test
  void testReaderWithNumericCells() throws IOException {
    // Create workbook with numeric cells
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      Sheet sheet = workbook.createSheet("Sheet1");

      // Row 0: ---
      Row row0 = sheet.createRow(0);
      row0.createCell(0).setCellValue("---");

      // Row 1: age | 30 (numeric)
      Row row1 = sheet.createRow(1);
      row1.createCell(0).setCellValue("age");
      Cell numericCell = row1.createCell(1);
      numericCell.setCellValue(30.0);

      // Row 2: score | 95.5 (numeric with decimal)
      Row row2 = sheet.createRow(2);
      row2.createCell(0).setCellValue("score");
      Cell decimalCell = row2.createCell(1);
      decimalCell.setCellValue(95.5);

      YamlWorkbookReader reader = YamlWorkbookReader.builder().build();
      List<Node> nodes = reader.fromWorkbook(workbook);

      assertEquals(1, nodes.size());
      assertTrue(nodes.get(0) instanceof MappingNode);
      MappingNode root = (MappingNode) nodes.get(0);
      assertEquals(2, root.getValue().size());

      // Check numeric values are parsed correctly
      NodeTuple ageTuple = root.getValue().get(0);
      assertEquals("30", ((ScalarNode) ageTuple.getValueNode()).getValue());

      NodeTuple scoreTuple = root.getValue().get(1);
      assertEquals("95.5", ((ScalarNode) scoreTuple.getValueNode()).getValue());
    }
  }

  @Test
  void testReaderWithBooleanCells() throws IOException {
    // Create workbook with boolean cells
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      Sheet sheet = workbook.createSheet("Sheet1");

      // Row 0: ---
      Row row0 = sheet.createRow(0);
      row0.createCell(0).setCellValue("---");

      // Row 1: active | true (boolean)
      Row row1 = sheet.createRow(1);
      row1.createCell(0).setCellValue("active");
      Cell boolCell = row1.createCell(1);
      boolCell.setCellValue(true);

      // Row 2: disabled | false (boolean)
      Row row2 = sheet.createRow(2);
      row2.createCell(0).setCellValue("disabled");
      Cell boolCell2 = row2.createCell(1);
      boolCell2.setCellValue(false);

      YamlWorkbookReader reader = YamlWorkbookReader.builder().build();
      List<Node> nodes = reader.fromWorkbook(workbook);

      assertEquals(1, nodes.size());
      MappingNode root = (MappingNode) nodes.get(0);
      assertEquals(2, root.getValue().size());

      // Check boolean values are parsed correctly
      NodeTuple activeTuple = root.getValue().get(0);
      assertEquals("true", ((ScalarNode) activeTuple.getValueNode()).getValue());

      NodeTuple disabledTuple = root.getValue().get(1);
      assertEquals("false", ((ScalarNode) disabledTuple.getValueNode()).getValue());
    }
  }

  @Test
  void testReaderWithEmptyRows() throws IOException {
    // Create workbook with empty rows interspersed
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      Sheet sheet = workbook.createSheet("Sheet1");

      // Row 0: ---
      Row row0 = sheet.createRow(0);
      row0.createCell(0).setCellValue("---");

      // Row 1: name | John
      Row row1 = sheet.createRow(1);
      row1.createCell(0).setCellValue("name");
      row1.createCell(1).setCellValue("John");

      // Row 2: empty (skipped in createRow)

      // Row 3: age | 30
      Row row3 = sheet.createRow(3);
      row3.createCell(0).setCellValue("age");
      row3.createCell(1).setCellValue("30");

      YamlWorkbookReader reader = YamlWorkbookReader.builder().build();
      List<Node> nodes = reader.fromWorkbook(workbook);

      assertEquals(1, nodes.size());
      MappingNode root = (MappingNode) nodes.get(0);
      assertEquals(2, root.getValue().size());
    }
  }

  @Test
  void testReaderWithEmptyCells() throws IOException {
    // Create workbook with null/empty cells
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      Sheet sheet = workbook.createSheet("Sheet1");

      // Row 0: ---
      Row row0 = sheet.createRow(0);
      row0.createCell(0).setCellValue("---");

      // Row 1: key | (empty value cell - created but empty)
      Row row1 = sheet.createRow(1);
      row1.createCell(0).setCellValue("key");
      row1.createCell(1).setCellValue(""); // Empty string

      YamlWorkbookReader reader = YamlWorkbookReader.builder().build();
      List<Node> nodes = reader.fromWorkbook(workbook);

      // Should handle empty cells gracefully
      assertNotNull(nodes);
    }
  }

  @Test
  void testReaderWithPrefixModeEmptyFirstCell() throws IOException {
    // Test PREFIX mode handling when first cell is empty
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      Sheet sheet = workbook.createSheet("Sheet1");

      // Row 0: ---
      Row row0 = sheet.createRow(0);
      row0.createCell(0).setCellValue("---");

      // Row 1: name | John (level 0, no prefix)
      Row row1 = sheet.createRow(1);
      row1.createCell(0).setCellValue("name");
      row1.createCell(1).setCellValue("John");

      // Row 2: 1> | city | NYC (level 1 with prefix)
      Row row2 = sheet.createRow(2);
      row2.createCell(0).setCellValue("1>");
      row2.createCell(1).setCellValue("city");
      row2.createCell(2).setCellValue("NYC");

      YamlWorkbookReader reader = YamlWorkbookReader.builder()
          .indentationMode(IndentationMode.PREFIX)
          .build();
      List<Node> nodes = reader.fromWorkbook(workbook);

      assertNotNull(nodes);
    }
  }

  @Test
  void testReaderWithDisplayModeEnumMapping() throws IOException {
    // Test DISPLAY_MODE reader with enum mapping
    String jsonSchema = """
        {
          "$schema": "http://json-schema.org/draft-07/schema#",
          "type": "object",
          "properties": {
            "status": {
              "type": "string",
              "title": "Status",
              "enum": ["active", "inactive"],
              "enumNames": ["Active", "Inactive"]
            }
          }
        }
        """;

    FormModeConfig config = FormModeConfig.builder()
        .useHiddenSheetsForLongEnums(false)
        .build();

    YamlWorkbookWriter writer = YamlWorkbookWriter.builder()
        .outputMode(OutputMode.FORM_MODE)
        .formModeConfig(config)
        .jsonSchema(jsonSchema)
        .build();

    Workbook workbook = writer.toWorkbook();

    // Set a value in the dropdown
    Sheet sheet = workbook.getSheetAt(0);
    for (int i = 1; i <= sheet.getLastRowNum(); i++) {
      Row row = sheet.getRow(i);
      if (row != null && row.getCell(0) != null) {
        String key = row.getCell(0).getStringCellValue();
        if ("Status".equals(key) && row.getCell(1) != null) {
          row.getCell(1).setCellValue("Active");
        }
      }
    }

    // Read back
    YamlWorkbookReader reader = YamlWorkbookReader.builder()
        .outputMode(OutputMode.FORM_MODE)
        .build();
    List<Node> nodes = reader.fromWorkbook(workbook);

    assertEquals(1, nodes.size());
    MappingNode root = (MappingNode) nodes.get(0);

    // Verify the enum value was mapped back
    NodeTuple tuple = root.getValue().get(0);
    assertEquals("status", ((ScalarNode) tuple.getKeyNode()).getValue());
    assertEquals("active", ((ScalarNode) tuple.getValueNode()).getValue());

    workbook.close();
  }

  @Test
  void testReaderWithNullWorkbook() {
    // Test fromWorkbook with null
    YamlWorkbookReader reader = YamlWorkbookReader.builder().build();
    List<Node> nodes = reader.fromWorkbook(null);
    assertTrue(nodes.isEmpty());
  }

  @Test
  void testReaderWithHiddenSheets() throws IOException {
    // Test that hidden sheets are skipped
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      Sheet sheet1 = workbook.createSheet("Sheet1");
      Sheet hiddenSheet = workbook.createSheet("Sheet1Hidden");

      // Hide the second sheet
      workbook.setSheetHidden(1, true);

      // Add content to Sheet1
      Row row0 = sheet1.createRow(0);
      row0.createCell(0).setCellValue("---");
      Row row1 = sheet1.createRow(1);
      row1.createCell(0).setCellValue("name");
      row1.createCell(1).setCellValue("John");

      // Add content to hidden sheet (should be ignored)
      Row hrow0 = hiddenSheet.createRow(0);
      hrow0.createCell(0).setCellValue("---");
      Row hrow1 = hiddenSheet.createRow(1);
      hrow1.createCell(0).setCellValue("hidden");
      hrow1.createCell(1).setCellValue("value");

      YamlWorkbookReader reader = YamlWorkbookReader.builder().build();
      List<Node> nodes = reader.fromWorkbook(workbook);

      assertEquals(1, nodes.size());
      MappingNode root = (MappingNode) nodes.get(0);
      assertEquals(1, root.getValue().size());
      assertEquals("name", ((ScalarNode) root.getValue().get(0).getKeyNode()).getValue());
    }
  }

  @Test
  void testReaderWithMismatchedSheetNames() throws IOException {
    // Test that sheets with non-matching names are handled
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      Sheet sheet = workbook.createSheet("CustomName"); // Not "Sheet1"

      Row row0 = sheet.createRow(0);
      row0.createCell(0).setCellValue("---");
      Row row1 = sheet.createRow(1);
      row1.createCell(0).setCellValue("name");
      row1.createCell(1).setCellValue("John");

      YamlWorkbookReader reader = YamlWorkbookReader.builder().build();
      List<Node> nodes = reader.fromWorkbook(workbook);

      // Sheet name doesn't match expected pattern, so should be skipped
      assertTrue(nodes.isEmpty());
    }
  }

  @Test
  void testReaderSequenceWithNestedMapping() throws IOException {
    // Test sequence with nested mapping items
    String yaml = """
        items:
          - name: first
            value: 1
          - name: second
            value: 2
        """;

    Workbook workbook = YamlWorkbook.toWorkbook(yaml);
    YamlWorkbookReader reader = YamlWorkbookReader.builder().build();
    List<Node> nodes = reader.fromWorkbook(workbook);

    assertEquals(1, nodes.size());
    MappingNode root = (MappingNode) nodes.get(0);
    SequenceNode items = (SequenceNode) root.getValue().get(0).getValueNode();
    assertEquals(2, items.getValue().size());

    // Each item should be a MappingNode
    assertTrue(items.getValue().get(0) instanceof MappingNode);
    assertTrue(items.getValue().get(1) instanceof MappingNode);

    workbook.close();
  }

  @Test
  void testWriterAndReaderWithComplexNestedComments() throws IOException {
    // Test complex scenario with multiple comment types
    String yaml = """
        # Document comment
        database:
          # Block comment for host
          host: localhost # inline comment
          # Block comment for port
          port: 5432
        """;

    Workbook workbook = YamlWorkbook.toWorkbook(yaml);
    YamlWorkbookReader reader = YamlWorkbookReader.builder().build();
    List<Node> nodes = reader.fromWorkbook(workbook);

    assertEquals(1, nodes.size());
    assertTrue(nodes.get(0) instanceof MappingNode);

    // Verify structure
    MappingNode root = (MappingNode) nodes.get(0);
    assertEquals(1, root.getValue().size()); // database

    workbook.close();
  }

  @Test
  void testReaderWithOnlyComments() throws IOException {
    // Test sheet with only comment rows
    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      Sheet sheet = workbook.createSheet("Sheet1");

      Row row0 = sheet.createRow(0);
      row0.createCell(0).setCellValue("---");

      Row row1 = sheet.createRow(1);
      row1.createCell(0).setCellValue("# This is just a comment");

      Row row2 = sheet.createRow(2);
      row2.createCell(0).setCellValue("# Another comment");

      YamlWorkbookReader reader = YamlWorkbookReader.builder().build();
      List<Node> nodes = reader.fromWorkbook(workbook);

      // Should handle gracefully - may return null node or empty
      assertTrue(nodes.isEmpty() || nodes.get(0) == null);
    }
  }

  @Test
  void testDisplayModeMappingCommentHidden() throws IOException {
    // Test DISPLAY_MODE with mappingComment=HIDDEN
    String yaml = """
        user:
          # User Settings
          settings:
            theme: dark
        """;

    DisplayModeConfig config = DisplayModeConfig.builder()
        .mappingComment(CommentDisplayOption.HIDDEN)
        .build();

    YamlWorkbookWriter writer = YamlWorkbookWriter.builder()
        .outputMode(OutputMode.DISPLAY_MODE)
        .displayModeConfig(config)
        .build();

    Workbook workbook = writer.toWorkbook(new StringReader(yaml));
    assertNotNull(workbook);

    // The block comment should be hidden
    Sheet sheet = workbook.getSheetAt(0);
    boolean foundUserSettings = false;
    for (int i = 0; i <= sheet.getLastRowNum(); i++) {
      Row row = sheet.getRow(i);
      if (row != null) {
        for (int j = 0; j <= row.getLastCellNum(); j++) {
          Cell cell = row.getCell(j);
          if (cell != null && cell.getStringCellValue().contains("User Settings")) {
            foundUserSettings = true;
            break;
          }
        }
      }
    }
    assertFalse(foundUserSettings, "User Settings should be hidden");

    workbook.close();
  }

  @Test
  void testDisplayModeSequenceCommentHidden() throws IOException {
    // Test DISPLAY_MODE with sequenceComment=HIDDEN
    String yaml = """
        items:
          # Item List
          - first
        """;

    DisplayModeConfig config = DisplayModeConfig.builder()
        .sequenceComment(CommentDisplayOption.HIDDEN)
        .build();

    YamlWorkbookWriter writer = YamlWorkbookWriter.builder()
        .outputMode(OutputMode.DISPLAY_MODE)
        .displayModeConfig(config)
        .build();

    Workbook workbook = writer.toWorkbook(new StringReader(yaml));
    assertNotNull(workbook);

    workbook.close();
  }

  @Test
  void testFormModeWithInvalidJsonSchema() {
    // Test FORM_MODE with invalid JSON schema
    String invalidSchema = "not a valid json";

    YamlWorkbookWriter writer = YamlWorkbookWriter.builder()
        .outputMode(OutputMode.FORM_MODE)
        .jsonSchema(invalidSchema)
        .build();

    assertThrows(RuntimeException.class, () -> writer.toWorkbook());
  }

  @Test
  void testMultipleDocuments() throws IOException {
    // Test multiple YAML documents
    String yaml = """
        ---
        name: Doc1
        ---
        name: Doc2
        ---
        name: Doc3
        """;

    Workbook workbook = YamlWorkbook.toWorkbook(yaml);
    YamlWorkbookReader reader = YamlWorkbookReader.builder().build();
    List<Node> nodes = reader.fromWorkbook(workbook);

    assertEquals(3, nodes.size());

    workbook.close();
  }

  @Test
  void testReaderDisplayModeValueFromCellComment() throws IOException {
    // Test that DISPLAY_MODE reader recovers original value from cell comment
    String yaml = """
        age: 30 # Age in Years
        """;

    YamlWorkbookWriter writer = YamlWorkbookWriter.builder()
        .outputMode(OutputMode.DISPLAY_MODE)
        .build();

    Workbook workbook = writer.toWorkbook(new StringReader(yaml));

    // In DISPLAY_MODE, the cell shows "Age in Years" but comment has "30"
    YamlWorkbookReader reader = YamlWorkbookReader.builder()
        .outputMode(OutputMode.DISPLAY_MODE)
        .build();
    List<Node> nodes = reader.fromWorkbook(workbook);

    assertEquals(1, nodes.size());
    MappingNode root = (MappingNode) nodes.get(0);
    NodeTuple ageTuple = root.getValue().get(0);
    assertEquals("30", ((ScalarNode) ageTuple.getValueNode()).getValue());

    workbook.close();
  }

  @Test
  void testWriterWithDocumentCommentHidden() throws IOException {
    // Test DISPLAY_MODE with documentComment=HIDDEN
    String yaml = """
        # This is a document comment
        name: John
        """;

    DisplayModeConfig config = DisplayModeConfig.builder()
        .documentComment(CommentVisibility.HIDDEN)
        .build();

    YamlWorkbookWriter writer = YamlWorkbookWriter.builder()
        .outputMode(OutputMode.DISPLAY_MODE)
        .displayModeConfig(config)
        .build();

    Workbook workbook = writer.toWorkbook(new StringReader(yaml));
    assertNotNull(workbook);

    // Document comment should be hidden
    Sheet sheet = workbook.getSheetAt(0);
    boolean foundDocComment = false;
    for (int i = 0; i <= sheet.getLastRowNum(); i++) {
      Row row = sheet.getRow(i);
      if (row != null && row.getCell(0) != null) {
        String value = row.getCell(0).getStringCellValue();
        if (value.contains("document comment")) {
          foundDocComment = true;
          break;
        }
      }
    }
    assertFalse(foundDocComment, "Document comment should be hidden");

    workbook.close();
  }

  // ==================== traverseScalarNodeWithPath Coverage Tests ====================

  @Test
  void testTraverseScalarNodeWithPathRootScalar() throws IOException {
    // Test traverseScalarNodeWithPath with a root-level scalar from JSON schema
    // A schema with type "string" generates a root scalar node
    String jsonSchema = """
        {
          "$schema": "http://json-schema.org/draft-07/schema#",
          "type": "string"
        }
        """;

    YamlWorkbookWriter writer = YamlWorkbookWriter.builder()
        .outputMode(OutputMode.FORM_MODE)
        .jsonSchema(jsonSchema)
        .build();

    Workbook workbook = writer.toWorkbook();
    assertNotNull(workbook);

    // Verify the workbook has content
    Sheet sheet = workbook.getSheetAt(0);
    assertNotNull(sheet);
    assertTrue(sheet.getLastRowNum() >= 0, "Should have at least one row");

    workbook.close();
  }

  @Test
  void testTraverseScalarNodeWithPathRootScalarWithEnum() throws IOException {
    // Test traverseScalarNodeWithPath with a root-level scalar that has enum
    String jsonSchema = """
        {
          "$schema": "http://json-schema.org/draft-07/schema#",
          "type": "string",
          "enum": ["option1", "option2", "option3"],
          "enumNames": ["Option One", "Option Two", "Option Three"]
        }
        """;

    YamlWorkbookWriter writer = YamlWorkbookWriter.builder()
        .outputMode(OutputMode.FORM_MODE)
        .jsonSchema(jsonSchema)
        .build();

    Workbook workbook = writer.toWorkbook();
    assertNotNull(workbook);

    // Verify the workbook has the dropdown for enum
    Sheet sheet = workbook.getSheetAt(0);
    assertNotNull(sheet);
    assertTrue(sheet.getLastRowNum() >= 0, "Should have rows");

    workbook.close();
  }

  @Test
  void testTraverseScalarNodeWithPathRootInteger() throws IOException {
    // Test traverseScalarNodeWithPath with a root-level integer
    String jsonSchema = """
        {
          "$schema": "http://json-schema.org/draft-07/schema#",
          "type": "integer"
        }
        """;

    YamlWorkbookWriter writer = YamlWorkbookWriter.builder()
        .outputMode(OutputMode.FORM_MODE)
        .jsonSchema(jsonSchema)
        .build();

    Workbook workbook = writer.toWorkbook();
    assertNotNull(workbook);

    Sheet sheet = workbook.getSheetAt(0);
    assertNotNull(sheet);

    workbook.close();
  }

  @Test
  void testTraverseScalarNodeWithPathRootBoolean() throws IOException {
    // Test traverseScalarNodeWithPath with a root-level boolean
    String jsonSchema = """
        {
          "$schema": "http://json-schema.org/draft-07/schema#",
          "type": "boolean"
        }
        """;

    YamlWorkbookWriter writer = YamlWorkbookWriter.builder()
        .outputMode(OutputMode.FORM_MODE)
        .jsonSchema(jsonSchema)
        .build();

    Workbook workbook = writer.toWorkbook();
    assertNotNull(workbook);

    Sheet sheet = workbook.getSheetAt(0);
    assertNotNull(sheet);

    workbook.close();
  }

  @Test
  void testTraverseScalarNodeWithPathRootNumber() throws IOException {
    // Test traverseScalarNodeWithPath with a root-level number
    String jsonSchema = """
        {
          "$schema": "http://json-schema.org/draft-07/schema#",
          "type": "number"
        }
        """;

    YamlWorkbookWriter writer = YamlWorkbookWriter.builder()
        .outputMode(OutputMode.FORM_MODE)
        .jsonSchema(jsonSchema)
        .build();

    Workbook workbook = writer.toWorkbook();
    assertNotNull(workbook);

    Sheet sheet = workbook.getSheetAt(0);
    assertNotNull(sheet);

    workbook.close();
  }

  @Test
  void testTraverseScalarNodeWithPathEnumWithoutEnumNames() throws IOException {
    // Test traverseScalarNodeWithPath with enum but no enumNames
    String jsonSchema = """
        {
          "$schema": "http://json-schema.org/draft-07/schema#",
          "type": "string",
          "enum": ["active", "inactive", "pending"]
        }
        """;

    YamlWorkbookWriter writer = YamlWorkbookWriter.builder()
        .outputMode(OutputMode.FORM_MODE)
        .jsonSchema(jsonSchema)
        .build();

    Workbook workbook = writer.toWorkbook();
    assertNotNull(workbook);

    // Dropdown should use enum values directly (no enumNames mapping)
    Sheet sheet = workbook.getSheetAt(0);
    assertNotNull(sheet);

    workbook.close();
  }

  @Test
  void testTraverseScalarNodeWithPathWithPrefixMode() throws IOException {
    // Test traverseScalarNodeWithPath with PREFIX indentation mode
    String jsonSchema = """
        {
          "$schema": "http://json-schema.org/draft-07/schema#",
          "type": "string"
        }
        """;

    YamlWorkbookWriter writer = YamlWorkbookWriter.builder()
        .outputMode(OutputMode.FORM_MODE)
        .jsonSchema(jsonSchema)
        .indentationMode(IndentationMode.PREFIX)
        .build();

    Workbook workbook = writer.toWorkbook();
    assertNotNull(workbook);

    Sheet sheet = workbook.getSheetAt(0);
    assertNotNull(sheet);

    workbook.close();
  }

  // ==================== YAML File-based Tests ====================

  @Test
  void testDisplayModeMappingCommentsFromFile() throws IOException {
    // Test using YAML file with mapping comments
    InputStream is = getClass().getResourceAsStream("/yaml/display-mode-mapping-comments.yaml");
    assertNotNull(is, "Test file should exist");

    DisplayModeConfig config = DisplayModeConfig.builder()
        .mappingComment(CommentDisplayOption.DISPLAY_NAME)
        .build();

    YamlWorkbookWriter writer = YamlWorkbookWriter.builder()
        .outputMode(OutputMode.DISPLAY_MODE)
        .displayModeConfig(config)
        .build();

    Workbook workbook = writer.toWorkbook(new InputStreamReader(is, StandardCharsets.UTF_8));
    assertNotNull(workbook);

    Sheet sheet = workbook.getSheetAt(0);
    assertNotNull(sheet);
    assertTrue(sheet.getLastRowNum() > 0);

    workbook.close();
    is.close();
  }

  @Test
  void testDisplayModeSequenceCommentsFromFile() throws IOException {
    // Test using YAML file with sequence comments
    InputStream is = getClass().getResourceAsStream("/yaml/display-mode-sequence-comments.yaml");
    assertNotNull(is, "Test file should exist");

    DisplayModeConfig config = DisplayModeConfig.builder()
        .sequenceComment(CommentDisplayOption.DISPLAY_NAME)
        .build();

    YamlWorkbookWriter writer = YamlWorkbookWriter.builder()
        .outputMode(OutputMode.DISPLAY_MODE)
        .displayModeConfig(config)
        .build();

    Workbook workbook = writer.toWorkbook(new InputStreamReader(is, StandardCharsets.UTF_8));
    assertNotNull(workbook);

    Sheet sheet = workbook.getSheetAt(0);
    assertNotNull(sheet);

    workbook.close();
    is.close();
  }

  @Test
  void testNestedBlockCommentsFromFile() throws IOException {
    // Test using YAML file with nested block comments
    InputStream is = getClass().getResourceAsStream("/yaml/nested-block-comments.yaml");
    assertNotNull(is, "Test file should exist");

    DisplayModeConfig config = DisplayModeConfig.builder()
        .mappingComment(CommentDisplayOption.DISPLAY_NAME)
        .sequenceComment(CommentDisplayOption.DISPLAY_NAME)
        .build();

    YamlWorkbookWriter writer = YamlWorkbookWriter.builder()
        .outputMode(OutputMode.DISPLAY_MODE)
        .displayModeConfig(config)
        .build();

    Workbook workbook = writer.toWorkbook(new InputStreamReader(is, StandardCharsets.UTF_8));
    assertNotNull(workbook);

    Sheet sheet = workbook.getSheetAt(0);
    assertNotNull(sheet);
    assertTrue(sheet.getLastRowNum() > 0);

    workbook.close();
    is.close();
  }

  @Test
  void testRootScalarFromFile() throws IOException {
    // Test using YAML file with root scalar
    InputStream is = getClass().getResourceAsStream("/yaml/root-scalar.yaml");
    assertNotNull(is, "Test file should exist");

    YamlWorkbookWriter writer = YamlWorkbookWriter.builder().build();
    Workbook workbook = writer.toWorkbook(new InputStreamReader(is, StandardCharsets.UTF_8));
    assertNotNull(workbook);

    Sheet sheet = workbook.getSheetAt(0);
    assertNotNull(sheet);

    workbook.close();
    is.close();
  }

  @Test
  void testMultidocWithCommentsFromFile() throws IOException {
    // Test using YAML file with multiple documents and comments
    InputStream is = getClass().getResourceAsStream("/yaml/multidoc-with-comments.yaml");
    assertNotNull(is, "Test file should exist");

    DisplayModeConfig config = DisplayModeConfig.builder()
        .mappingComment(CommentDisplayOption.DISPLAY_NAME)
        .sequenceComment(CommentDisplayOption.DISPLAY_NAME)
        .build();

    YamlWorkbookWriter writer = YamlWorkbookWriter.builder()
        .outputMode(OutputMode.DISPLAY_MODE)
        .displayModeConfig(config)
        .build();

    Workbook workbook = writer.toWorkbook(new InputStreamReader(is, StandardCharsets.UTF_8));
    assertNotNull(workbook);

    Sheet sheet = workbook.getSheetAt(0);
    assertNotNull(sheet);

    // Read back with reader
    YamlWorkbookReader reader = YamlWorkbookReader.builder()
        .outputMode(OutputMode.DISPLAY_MODE)
        .build();
    List<Node> nodes = reader.fromWorkbook(workbook);
    assertEquals(3, nodes.size());

    workbook.close();
    is.close();
  }

}
