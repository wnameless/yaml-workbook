package com.github.wnameless.workbook.yamlworkbook;

import static org.junit.jupiter.api.Assertions.*;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.jupiter.api.BeforeAll;
import org.junit.jupiter.api.Test;

class YamlWorkbookWriterTest {

  private static final boolean WRITE_EXCEL_FILES = true;
  private static final Path OUTPUT_DIR = Paths.get("target/test-excel");

  @BeforeAll
  static void setUpOnce() throws IOException {
    if (WRITE_EXCEL_FILES) {
      Files.createDirectories(OUTPUT_DIR);
    }
  }

  @Test
  void testSimpleYaml() throws IOException {
    String yaml = loadYaml("yaml/simple.yaml");
    Workbook workbook = YamlWorkbook.toWorkbook(yaml);

    assertNotNull(workbook);
    Sheet sheet = workbook.getSheetAt(0);
    assertNotNull(sheet);

    // --- (frontmatter)
    Row row0 = sheet.getRow(0);
    assertEquals("---", row0.getCell(0).getStringCellValue());

    // name: John Doe
    Row row1 = sheet.getRow(1);
    assertEquals("name", row1.getCell(0).getStringCellValue());
    assertEquals("John Doe", row1.getCell(1).getStringCellValue());

    // age: 30
    Row row2 = sheet.getRow(2);
    assertEquals("age", row2.getCell(0).getStringCellValue());
    assertEquals("30", row2.getCell(1).getStringCellValue());

    // city: New York
    Row row3 = sheet.getRow(3);
    assertEquals("city", row3.getCell(0).getStringCellValue());
    assertEquals("New York", row3.getCell(1).getStringCellValue());

    writeExcelFile(workbook, "simple.xlsx");
    workbook.close();
  }

  @Test
  void testNestedYaml() throws IOException {
    String yaml = loadYaml("yaml/nested.yaml");
    Workbook workbook = YamlWorkbook.toWorkbook(yaml);

    assertNotNull(workbook);
    Sheet sheet = workbook.getSheetAt(0);

    // --- (frontmatter)
    Row row0 = sheet.getRow(0);
    assertEquals("---", row0.getCell(0).getStringCellValue());

    // person:
    Row row1 = sheet.getRow(1);
    assertEquals("person", row1.getCell(0).getStringCellValue());

    // name: Jane Smith (indented)
    Row row2 = sheet.getRow(2);
    assertEquals("name", row2.getCell(1).getStringCellValue());
    assertEquals("Jane Smith", row2.getCell(2).getStringCellValue());

    // address:
    Row row3 = sheet.getRow(3);
    assertEquals("address", row3.getCell(1).getStringCellValue());

    // street: 123 Main St (double indented)
    Row row4 = sheet.getRow(4);
    assertEquals("street", row4.getCell(2).getStringCellValue());
    assertEquals("123 Main St", row4.getCell(3).getStringCellValue());

    writeExcelFile(workbook, "nested.xlsx");
    workbook.close();
  }

  @Test
  void testSequenceYaml() throws IOException {
    String yaml = loadYaml("yaml/sequence.yaml");
    Workbook workbook = YamlWorkbook.toWorkbook(yaml);

    assertNotNull(workbook);
    Sheet sheet = workbook.getSheetAt(0);

    // --- (frontmatter)
    Row row0 = sheet.getRow(0);
    assertEquals("---", row0.getCell(0).getStringCellValue());

    // fruits:
    Row row1 = sheet.getRow(1);
    assertEquals("fruits", row1.getCell(0).getStringCellValue());

    // - apple
    Row row2 = sheet.getRow(2);
    assertEquals("-", row2.getCell(1).getStringCellValue());
    assertEquals("apple", row2.getCell(2).getStringCellValue());

    // - banana
    Row row3 = sheet.getRow(3);
    assertEquals("-", row3.getCell(1).getStringCellValue());
    assertEquals("banana", row3.getCell(2).getStringCellValue());

    // - cherry
    Row row4 = sheet.getRow(4);
    assertEquals("-", row4.getCell(1).getStringCellValue());
    assertEquals("cherry", row4.getCell(2).getStringCellValue());

    writeExcelFile(workbook, "sequence.xlsx");
    workbook.close();
  }

  @Test
  void testCommentsYaml() throws IOException {
    String yaml = loadYaml("yaml/comments.yaml");
    Workbook workbook = YamlWorkbook.toWorkbook(yaml);

    assertNotNull(workbook);
    Sheet sheet = workbook.getSheetAt(0);

    int rowNum = 0;

    // Row 0: --- (frontmatter)
    Row row = sheet.getRow(rowNum++);
    assertEquals("---", row.getCell(0).getStringCellValue());

    // Row 1: # This is a configuration file (block comment)
    row = sheet.getRow(rowNum++);
    assertTrue(row.getCell(0).getStringCellValue().startsWith("#"));

    // Row 2: database: # main database config (key with inline comment, complex value)
    row = sheet.getRow(rowNum++);
    assertEquals("database", row.getCell(0).getStringCellValue());
    assertEquals("# main database config", row.getCell(1).getStringCellValue());

    // Row 3: # Database connection settings (block comment, indented)
    row = sheet.getRow(rowNum++);
    assertTrue(row.getCell(1).getStringCellValue().startsWith("#"));

    // Row 4: host: # server address localhost # primary host (key inline + value + value inline)
    row = sheet.getRow(rowNum++);
    assertEquals("host", row.getCell(1).getStringCellValue());
    assertEquals("# server address", row.getCell(2).getStringCellValue());
    assertEquals("localhost", row.getCell(3).getStringCellValue());
    assertEquals("# primary host", row.getCell(4).getStringCellValue());

    // Row 5: port: 5432 # default PostgreSQL port
    row = sheet.getRow(rowNum++);
    assertEquals("port", row.getCell(1).getStringCellValue());
    assertEquals("5432", row.getCell(2).getStringCellValue());
    assertEquals("# default PostgreSQL port", row.getCell(3).getStringCellValue());

    // Row 6: # Credentials (block comment)
    row = sheet.getRow(rowNum++);
    assertTrue(row.getCell(1).getStringCellValue().startsWith("#"));

    // Row 7: username: admin (no inline comment)
    row = sheet.getRow(rowNum++);
    assertEquals("username", row.getCell(1).getStringCellValue());
    assertEquals("admin", row.getCell(2).getStringCellValue());
    assertNull(row.getCell(3));

    // Row 8: password: secret # change in production
    row = sheet.getRow(rowNum++);
    assertEquals("password", row.getCell(1).getStringCellValue());
    assertEquals("secret", row.getCell(2).getStringCellValue());
    assertEquals("# change in production", row.getCell(3).getStringCellValue());

    // Row 9: connection: # connection pool settings (key with inline comment, complex value)
    row = sheet.getRow(rowNum++);
    assertEquals("connection", row.getCell(1).getStringCellValue());
    assertEquals("# connection pool settings", row.getCell(2).getStringCellValue());

    // Row 10: pool_size: # max connections 10 (key inline + value)
    row = sheet.getRow(rowNum++);
    assertEquals("pool_size", row.getCell(2).getStringCellValue());
    assertEquals("# max connections", row.getCell(3).getStringCellValue());
    assertEquals("10", row.getCell(4).getStringCellValue());

    // Row 11: timeout: # in seconds 30 # 30 seconds (key inline + value + value inline)
    row = sheet.getRow(rowNum++);
    assertEquals("timeout", row.getCell(2).getStringCellValue());
    assertEquals("# in seconds", row.getCell(3).getStringCellValue());
    assertEquals("30", row.getCell(4).getStringCellValue());
    assertEquals("# 30 seconds", row.getCell(5).getStringCellValue());

    // Row 12: replicas: # read replicas (key with inline comment, sequence value)
    row = sheet.getRow(rowNum++);
    assertEquals("replicas", row.getCell(1).getStringCellValue());
    assertEquals("# read replicas", row.getCell(2).getStringCellValue());

    // Row 13: - (sequence item mark for first replica)
    row = sheet.getRow(rowNum++);
    assertEquals("-", row.getCell(2).getStringCellValue());

    // Row 14: host: replica1.local # first replica
    row = sheet.getRow(rowNum++);
    assertEquals("host", row.getCell(3).getStringCellValue());
    assertEquals("replica1.local", row.getCell(4).getStringCellValue());
    assertEquals("# first replica", row.getCell(5).getStringCellValue());

    // Row 15: port: 5432
    row = sheet.getRow(rowNum++);
    assertEquals("port", row.getCell(3).getStringCellValue());
    assertEquals("5432", row.getCell(4).getStringCellValue());

    // Row 16: - (sequence item mark for second replica)
    row = sheet.getRow(rowNum++);
    assertEquals("-", row.getCell(2).getStringCellValue());

    // Row 17: host: # second replica address replica2.local (key inline + value)
    row = sheet.getRow(rowNum++);
    assertEquals("host", row.getCell(3).getStringCellValue());
    assertEquals("# second replica address", row.getCell(4).getStringCellValue());
    assertEquals("replica2.local", row.getCell(5).getStringCellValue());

    // Row 18: port: 5433 # non-standard port
    row = sheet.getRow(rowNum++);
    assertEquals("port", row.getCell(3).getStringCellValue());
    assertEquals("5433", row.getCell(4).getStringCellValue());
    assertEquals("# non-standard port", row.getCell(5).getStringCellValue());

    // Row 19: allowed_ips: # whitelist (key with inline comment, simple array)
    row = sheet.getRow(rowNum++);
    assertEquals("allowed_ips", row.getCell(1).getStringCellValue());
    assertEquals("# whitelist", row.getCell(2).getStringCellValue());

    // Row 20: - 192.168.1.1 # main server
    row = sheet.getRow(rowNum++);
    assertEquals("-", row.getCell(2).getStringCellValue());
    assertEquals("192.168.1.1", row.getCell(3).getStringCellValue());
    assertEquals("# main server", row.getCell(4).getStringCellValue());

    // Row 21: - 192.168.1.2 # backup server
    row = sheet.getRow(rowNum++);
    assertEquals("-", row.getCell(2).getStringCellValue());
    assertEquals("192.168.1.2", row.getCell(3).getStringCellValue());
    assertEquals("# backup server", row.getCell(4).getStringCellValue());

    // Row 22: # internal network (block comment before array item)
    row = sheet.getRow(rowNum++);
    assertTrue(row.getCell(2).getStringCellValue().startsWith("#"));

    // Row 23: - 10.0.0.1
    row = sheet.getRow(rowNum++);
    assertEquals("-", row.getCell(2).getStringCellValue());
    assertEquals("10.0.0.1", row.getCell(3).getStringCellValue());

    writeExcelFile(workbook, "comments.xlsx");
    workbook.close();
  }

  @Test
  void testComplexYaml() throws IOException {
    String yaml = loadYaml("yaml/complex.yaml");
    Workbook workbook = YamlWorkbook.toWorkbook(yaml);

    assertNotNull(workbook);
    Sheet sheet = workbook.getSheetAt(0);
    assertTrue(sheet.getLastRowNum() > 0);

    writeExcelFile(workbook, "complex.xlsx");
    workbook.close();
  }

  @Test
  void testEmptyYaml() throws IOException {
    Workbook workbook = YamlWorkbook.toWorkbook("");

    assertNotNull(workbook);
    Sheet sheet = workbook.getSheetAt(0);
    assertEquals(-1, sheet.getLastRowNum());

    writeExcelFile(workbook, "empty.xlsx");
    workbook.close();
  }

  @Test
  void testScalarOnlyYaml() throws IOException {
    Workbook workbook = YamlWorkbook.toWorkbook("hello world");

    assertNotNull(workbook);
    Sheet sheet = workbook.getSheetAt(0);

    // --- (frontmatter)
    Row row0 = sheet.getRow(0);
    assertEquals("---", row0.getCell(0).getStringCellValue());

    Row row1 = sheet.getRow(1);
    assertEquals("hello world", row1.getCell(0).getStringCellValue());

    writeExcelFile(workbook, "scalar.xlsx");
    workbook.close();
  }

  @Test
  void testMultiDocumentYaml() throws IOException {
    String yaml = loadYaml("yaml/multidoc.yaml");
    Workbook workbook = YamlWorkbook.toWorkbook(yaml);

    assertNotNull(workbook);
    Sheet sheet = workbook.getSheetAt(0);

    // --- (frontmatter for Document One)
    Row row0 = sheet.getRow(0);
    assertEquals("---", row0.getCell(0).getStringCellValue());

    // Document One: name, type
    Row row1 = sheet.getRow(1);
    assertEquals("name", row1.getCell(0).getStringCellValue());
    assertEquals("Document One", row1.getCell(1).getStringCellValue());

    Row row2 = sheet.getRow(2);
    assertEquals("type", row2.getCell(0).getStringCellValue());
    assertEquals("config", row2.getCell(1).getStringCellValue());

    // --- (frontmatter for Document Two)
    Row row3 = sheet.getRow(3);
    assertEquals("---", row3.getCell(0).getStringCellValue());

    // Document Two: name, type, items
    Row row4 = sheet.getRow(4);
    assertEquals("name", row4.getCell(0).getStringCellValue());
    assertEquals("Document Two", row4.getCell(1).getStringCellValue());

    Row row5 = sheet.getRow(5);
    assertEquals("type", row5.getCell(0).getStringCellValue());
    assertEquals("data", row5.getCell(1).getStringCellValue());

    Row row6 = sheet.getRow(6);
    assertEquals("items", row6.getCell(0).getStringCellValue());

    // - first
    Row row7 = sheet.getRow(7);
    assertEquals("-", row7.getCell(1).getStringCellValue());
    assertEquals("first", row7.getCell(2).getStringCellValue());

    // - second
    Row row8 = sheet.getRow(8);
    assertEquals("-", row8.getCell(1).getStringCellValue());
    assertEquals("second", row8.getCell(2).getStringCellValue());

    // --- (frontmatter for Document Three)
    Row row9 = sheet.getRow(9);
    assertEquals("---", row9.getCell(0).getStringCellValue());

    // Document Three: name, settings
    Row row10 = sheet.getRow(10);
    assertEquals("name", row10.getCell(0).getStringCellValue());
    assertEquals("Document Three", row10.getCell(1).getStringCellValue());

    Row row11 = sheet.getRow(11);
    assertEquals("settings", row11.getCell(0).getStringCellValue());

    Row row12 = sheet.getRow(12);
    assertEquals("enabled", row12.getCell(1).getStringCellValue());
    assertEquals("true", row12.getCell(2).getStringCellValue());

    Row row13 = sheet.getRow(13);
    assertEquals("level", row13.getCell(1).getStringCellValue());
    assertEquals("5", row13.getCell(2).getStringCellValue());

    writeExcelFile(workbook, "multidoc.xlsx");
    workbook.close();
  }

  @Test
  void testValueStartingWithHashEscaping() throws IOException {
    // Test that values starting with # are escaped with backslash
    String yaml = """
        message: '# Look like a comment'
        hashtag: '#trending'
        normal: regular value
        mixed: 'Hello # world'
        """;
    Workbook workbook = YamlWorkbook.toWorkbook(yaml);

    assertNotNull(workbook);
    Sheet sheet = workbook.getSheetAt(0);

    // --- (frontmatter)
    assertEquals("---", sheet.getRow(0).getCell(0).getStringCellValue());

    // message: \# Look like a comment (escaped because starts with #)
    Row row1 = sheet.getRow(1);
    assertEquals("message", row1.getCell(0).getStringCellValue());
    assertEquals("\\# Look like a comment", row1.getCell(1).getStringCellValue());

    // hashtag: \#trending (escaped because starts with #)
    Row row2 = sheet.getRow(2);
    assertEquals("hashtag", row2.getCell(0).getStringCellValue());
    assertEquals("\\#trending", row2.getCell(1).getStringCellValue());

    // normal: regular value (not escaped)
    Row row3 = sheet.getRow(3);
    assertEquals("normal", row3.getCell(0).getStringCellValue());
    assertEquals("regular value", row3.getCell(1).getStringCellValue());

    // mixed: Hello # world (not escaped because # is not at the start)
    Row row4 = sheet.getRow(4);
    assertEquals("mixed", row4.getCell(0).getStringCellValue());
    assertEquals("Hello # world", row4.getCell(1).getStringCellValue());

    writeExcelFile(workbook, "escaped_hash.xlsx");
    workbook.close();
  }

  @Test
  void testSequenceItemStartingWithHashEscaping() throws IOException {
    // Test sequence items that start with # are escaped
    String yaml = """
        items:
          - '# first item'
          - normal item
          - '# another hash item'
        """;
    Workbook workbook = YamlWorkbook.toWorkbook(yaml);

    Sheet sheet = workbook.getSheetAt(0);

    // --- (frontmatter)
    assertEquals("---", sheet.getRow(0).getCell(0).getStringCellValue());

    // items:
    assertEquals("items", sheet.getRow(1).getCell(0).getStringCellValue());

    // - \# first item (escaped)
    Row row2 = sheet.getRow(2);
    assertEquals("-", row2.getCell(1).getStringCellValue());
    assertEquals("\\# first item", row2.getCell(2).getStringCellValue());

    // - normal item (not escaped)
    Row row3 = sheet.getRow(3);
    assertEquals("-", row3.getCell(1).getStringCellValue());
    assertEquals("normal item", row3.getCell(2).getStringCellValue());

    // - \# another hash item (escaped)
    Row row4 = sheet.getRow(4);
    assertEquals("-", row4.getCell(1).getStringCellValue());
    assertEquals("\\# another hash item", row4.getCell(2).getStringCellValue());

    writeExcelFile(workbook, "escaped_sequence.xlsx");
    workbook.close();
  }

  @Test
  void testValueStartingWithBackslashEscaping() throws IOException {
    // Test that values starting with \ are double-escaped
    String yaml = """
        path: '\\some\\path'
        escaped_hash: '\\# not a comment'
        """;
    Workbook workbook = YamlWorkbook.toWorkbook(yaml);

    Sheet sheet = workbook.getSheetAt(0);

    // path: \\some\path (only first \ is escaped because it's at the start)
    Row row1 = sheet.getRow(1);
    assertEquals("path", row1.getCell(0).getStringCellValue());
    assertEquals("\\\\some\\path", row1.getCell(1).getStringCellValue());

    // escaped_hash: \\# not a comment (double-escaped)
    Row row2 = sheet.getRow(2);
    assertEquals("escaped_hash", row2.getCell(0).getStringCellValue());
    assertEquals("\\\\# not a comment", row2.getCell(1).getStringCellValue());

    writeExcelFile(workbook, "escaped_backslash.xlsx");
    workbook.close();
  }

  // ==================== WORKBOOK_READABLE Mode Tests ====================

  @Test
  void testDisplayModeDefaultConfig() throws IOException {
    // Test with simple key-value pairs and inline value comments
    String yaml = """
        name: John Doe
        email: john@example.com
        age: 30 # Age in Years
        """;

    YamlWorkbookWriter writer = YamlWorkbookWriter.builder()
        .outputMode(OutputMode.DISPLAY_MODE)
        .build();

    Workbook workbook = writer.toWorkbook(new java.io.StringReader(yaml));

    assertNotNull(workbook);
    Sheet sheet = workbook.getSheetAt(0);

    writeExcelFile(workbook, "display_mode_default.xlsx");

    int rowNum = 0;

    // Row 0: --- (frontmatter)
    Row row = sheet.getRow(rowNum++);
    assertEquals("---", row.getCell(0).getStringCellValue());

    // Row 1: name | John Doe
    row = sheet.getRow(rowNum++);
    assertEquals("name", row.getCell(0).getStringCellValue());
    assertEquals("John Doe", row.getCell(1).getStringCellValue());

    // Row 2: email | john@example.com
    row = sheet.getRow(rowNum++);
    assertEquals("email", row.getCell(0).getStringCellValue());
    assertEquals("john@example.com", row.getCell(1).getStringCellValue());

    // Row 3: age | Age in Years (value replaced by inline comment with DISPLAY_NAME)
    row = sheet.getRow(rowNum++);
    assertEquals("age", row.getCell(0).getStringCellValue());
    assertEquals("Age in Years", row.getCell(1).getStringCellValue());

    workbook.close();
  }

  @Test
  void testDisplayModeKeyValuePairCommentVisible() throws IOException {
    // Test with KEY_VALUE_PAIR comments set to COMMENT (visible)
    String yaml = """
        # User Information
        name: John
        # Contact Details
        email: john@example.com
        """;

    DisplayModeConfig config = DisplayModeConfig.builder()
        .keyValuePairComment(CommentVisibility.COMMENT)
        .build();

    YamlWorkbookWriter writer = YamlWorkbookWriter.builder()
        .outputMode(OutputMode.DISPLAY_MODE)
        .displayModeConfig(config)
        .build();

    Workbook workbook = writer.toWorkbook(new java.io.StringReader(yaml));
    Sheet sheet = workbook.getSheetAt(0);

    int rowNum = 0;

    // Row 0: ---
    assertEquals("---", sheet.getRow(rowNum++).getCell(0).getStringCellValue());

    // Row 1: # User Information (KEY_VALUE_PAIR comment shown)
    Row row1 = sheet.getRow(rowNum++);
    assertTrue(row1.getCell(0).getStringCellValue().startsWith("#"));

    // Row 2: name | John
    Row row2 = sheet.getRow(rowNum++);
    assertEquals("name", row2.getCell(0).getStringCellValue());
    assertEquals("John", row2.getCell(1).getStringCellValue());

    // Row 3: # Contact Details (KEY_VALUE_PAIR comment shown)
    Row row3 = sheet.getRow(rowNum++);
    assertTrue(row3.getCell(0).getStringCellValue().startsWith("#"));

    // Row 4: email | john@example.com
    Row row4 = sheet.getRow(rowNum++);
    assertEquals("email", row4.getCell(0).getStringCellValue());
    assertEquals("john@example.com", row4.getCell(1).getStringCellValue());

    writeExcelFile(workbook, "display_mode_kvp_comment.xlsx");
    workbook.close();
  }

  @Test
  void testDisplayModeKeyValuePairCommentHidden() throws IOException {
    // Test with KEY_VALUE_PAIR comments set to HIDDEN (default)
    String yaml = """
        # User Information
        name: John
        # Contact Details
        email: john@example.com
        """;

    YamlWorkbookWriter writer = YamlWorkbookWriter.builder()
        .outputMode(OutputMode.DISPLAY_MODE)
        .build();

    Workbook workbook = writer.toWorkbook(new java.io.StringReader(yaml));
    Sheet sheet = workbook.getSheetAt(0);

    int rowNum = 0;

    // Row 0: ---
    assertEquals("---", sheet.getRow(rowNum++).getCell(0).getStringCellValue());

    // Row 1: name | John (comment hidden)
    Row row1 = sheet.getRow(rowNum++);
    assertEquals("name", row1.getCell(0).getStringCellValue());
    assertEquals("John", row1.getCell(1).getStringCellValue());

    // Row 2: email | john@example.com (comment hidden)
    Row row2 = sheet.getRow(rowNum++);
    assertEquals("email", row2.getCell(0).getStringCellValue());
    assertEquals("john@example.com", row2.getCell(1).getStringCellValue());

    writeExcelFile(workbook, "display_mode_kvp_hidden.xlsx");
    workbook.close();
  }

  @Test
  void testDisplayModeDocumentCommentVisible() throws IOException {
    // Test with DOCUMENT comment set to COMMENT (visible)
    // Note: In SnakeYAML, comments before the first key are block comments of the root mapping node
    String yaml = """
        # This is a document title
        name: John
        """;

    DisplayModeConfig config = DisplayModeConfig.builder()
        .documentComment(CommentVisibility.COMMENT)
        .keyValuePairComment(CommentVisibility.HIDDEN)  // Hide key-value pair comments
        .build();

    YamlWorkbookWriter writer = YamlWorkbookWriter.builder()
        .outputMode(OutputMode.DISPLAY_MODE)
        .displayModeConfig(config)
        .build();

    Workbook workbook = writer.toWorkbook(new java.io.StringReader(yaml));
    Sheet sheet = workbook.getSheetAt(0);

    // Write file for inspection
    writeExcelFile(workbook, "display_mode_doc_comment.xlsx");

    int rowNum = 0;

    // Row 0: ---
    assertEquals("---", sheet.getRow(rowNum++).getCell(0).getStringCellValue());

    // Row 1: name | John (comments before first key are on key node as block comments, not document)
    Row row1 = sheet.getRow(rowNum++);
    assertEquals("name", row1.getCell(0).getStringCellValue());
    assertEquals("John", row1.getCell(1).getStringCellValue());

    workbook.close();
  }

  @Test
  void testDisplayModeItemCommentVisible() throws IOException {
    // Test with ITEM comment set to COMMENT (visible)
    String yaml = """
        fruits:
          # A red fruit
          - apple
          # A yellow fruit
          - banana
        """;

    DisplayModeConfig config = DisplayModeConfig.builder()
        .itemComment(CommentVisibility.COMMENT)
        .build();

    YamlWorkbookWriter writer = YamlWorkbookWriter.builder()
        .outputMode(OutputMode.DISPLAY_MODE)
        .displayModeConfig(config)
        .build();

    Workbook workbook = writer.toWorkbook(new java.io.StringReader(yaml));
    Sheet sheet = workbook.getSheetAt(0);

    int rowNum = 0;

    // Row 0: ---
    assertEquals("---", sheet.getRow(rowNum++).getCell(0).getStringCellValue());

    // Row 1: fruits
    assertEquals("fruits", sheet.getRow(rowNum++).getCell(0).getStringCellValue());

    // Row 2: # A red fruit (item comment shown)
    Row row2 = sheet.getRow(rowNum++);
    assertTrue(row2.getCell(1).getStringCellValue().startsWith("#"));

    // Row 3: - | apple
    Row row3 = sheet.getRow(rowNum++);
    assertEquals("-", row3.getCell(1).getStringCellValue());
    assertEquals("apple", row3.getCell(2).getStringCellValue());

    // Row 4: # A yellow fruit (item comment shown)
    Row row4 = sheet.getRow(rowNum++);
    assertTrue(row4.getCell(1).getStringCellValue().startsWith("#"));

    // Row 5: - | banana
    Row row5 = sheet.getRow(rowNum++);
    assertEquals("-", row5.getCell(1).getStringCellValue());
    assertEquals("banana", row5.getCell(2).getStringCellValue());

    writeExcelFile(workbook, "display_mode_item_comment.xlsx");
    workbook.close();
  }

  @Test
  void testDisplayModeItemCommentHidden() throws IOException {
    // Test with ITEM comment set to HIDDEN (default)
    String yaml = """
        fruits:
          # A red fruit
          - apple
          # A yellow fruit
          - banana
        """;

    YamlWorkbookWriter writer = YamlWorkbookWriter.builder()
        .outputMode(OutputMode.DISPLAY_MODE)
        .build();

    Workbook workbook = writer.toWorkbook(new java.io.StringReader(yaml));
    Sheet sheet = workbook.getSheetAt(0);

    int rowNum = 0;

    // Row 0: ---
    assertEquals("---", sheet.getRow(rowNum++).getCell(0).getStringCellValue());

    // Row 1: fruits
    assertEquals("fruits", sheet.getRow(rowNum++).getCell(0).getStringCellValue());

    // Row 2: - | apple (comment hidden)
    Row row2 = sheet.getRow(rowNum++);
    assertEquals("-", row2.getCell(1).getStringCellValue());
    assertEquals("apple", row2.getCell(2).getStringCellValue());

    // Row 3: - | banana (comment hidden)
    Row row3 = sheet.getRow(rowNum++);
    assertEquals("-", row3.getCell(1).getStringCellValue());
    assertEquals("banana", row3.getCell(2).getStringCellValue());

    writeExcelFile(workbook, "display_mode_item_hidden.xlsx");
    workbook.close();
  }

  @Test
  void testDisplayModeObjectComment() throws IOException {
    // Test OBJECT comment - when a nested mapping has a block comment before it
    // Block comments before keys whose values are MappingNodes use mappingComment config
    String yaml = """
        user:
          # User Settings
          settings:
            theme: dark
        """;

    // Test DISPLAY_NAME (default) - the block comment before "settings" key is shown as header
    YamlWorkbookWriter writer = YamlWorkbookWriter.builder()
        .outputMode(OutputMode.DISPLAY_MODE)
        .build();

    Workbook workbook = writer.toWorkbook(new java.io.StringReader(yaml));
    Sheet sheet = workbook.getSheetAt(0);

    writeExcelFile(workbook, "display_mode_object_displayname.xlsx");

    int rowNum = 0;

    // Row 0: ---
    assertEquals("---", sheet.getRow(rowNum++).getCell(0).getStringCellValue());

    // Row 1: user
    assertEquals("user", sheet.getRow(rowNum++).getCell(0).getStringCellValue());

    // Row 2: User Settings (header row from mappingComment=DISPLAY_NAME)
    assertEquals("User Settings", sheet.getRow(rowNum++).getCell(1).getStringCellValue());

    // Row 3: settings
    assertEquals("settings", sheet.getRow(rowNum++).getCell(1).getStringCellValue());

    // Row 4: theme | dark
    Row row4 = sheet.getRow(rowNum++);
    assertEquals("theme", row4.getCell(2).getStringCellValue());
    assertEquals("dark", row4.getCell(3).getStringCellValue());

    workbook.close();
  }

  @Test
  void testDisplayModeObjectCommentHidden() throws IOException {
    // Test that object comment HIDDEN works (same as default since KEY_VALUE_PAIR is HIDDEN)
    String yaml = """
        # User Profile
        user:
          name: John
        """;

    DisplayModeConfig config = DisplayModeConfig.builder()
        .mappingComment(CommentDisplayOption.HIDDEN)
        .keyValuePairComment(CommentVisibility.HIDDEN)
        .build();

    YamlWorkbookWriter writer = YamlWorkbookWriter.builder()
        .outputMode(OutputMode.DISPLAY_MODE)
        .displayModeConfig(config)
        .build();

    Workbook workbook = writer.toWorkbook(new java.io.StringReader(yaml));
    Sheet sheet = workbook.getSheetAt(0);

    writeExcelFile(workbook, "display_mode_object_hidden.xlsx");

    int rowNum = 0;

    // Row 0: ---
    assertEquals("---", sheet.getRow(rowNum++).getCell(0).getStringCellValue());

    // Row 1: user (comment "User Profile" hidden - it's a KEY_VALUE_PAIR comment)
    assertEquals("user", sheet.getRow(rowNum++).getCell(0).getStringCellValue());

    // Row 2: name | John
    Row row2 = sheet.getRow(rowNum++);
    assertEquals("name", row2.getCell(1).getStringCellValue());
    assertEquals("John", row2.getCell(2).getStringCellValue());

    workbook.close();
  }

  @Test
  void testDisplayModeArrayComment() throws IOException {
    // Test ARRAY comment - block comments before keys whose values are SequenceNodes
    // use sequenceComment config
    String yaml = """
        # Fruit List
        fruits:
          - apple
          - banana
        """;

    YamlWorkbookWriter writer = YamlWorkbookWriter.builder()
        .outputMode(OutputMode.DISPLAY_MODE)
        .build();

    Workbook workbook = writer.toWorkbook(new java.io.StringReader(yaml));
    Sheet sheet = workbook.getSheetAt(0);

    writeExcelFile(workbook, "display_mode_array_displayname.xlsx");

    int rowNum = 0;

    // Row 0: ---
    assertEquals("---", sheet.getRow(rowNum++).getCell(0).getStringCellValue());

    // Row 1: Fruit List (header row from sequenceComment=DISPLAY_NAME)
    assertEquals("Fruit List", sheet.getRow(rowNum++).getCell(0).getStringCellValue());

    // Row 2: fruits
    assertEquals("fruits", sheet.getRow(rowNum++).getCell(0).getStringCellValue());

    // Row 3: - | apple
    Row row3 = sheet.getRow(rowNum++);
    assertEquals("-", row3.getCell(1).getStringCellValue());
    assertEquals("apple", row3.getCell(2).getStringCellValue());

    // Row 4: - | banana
    Row row4 = sheet.getRow(rowNum++);
    assertEquals("-", row4.getCell(1).getStringCellValue());
    assertEquals("banana", row4.getCell(2).getStringCellValue());

    workbook.close();
  }

  @Test
  void testDisplayModeAllCommentsAsComment() throws IOException {
    // Test with KEY_VALUE_PAIR comments set to COMMENT (show as comments)
    String yaml = """
        # First comment
        # Second comment
        name: John
        """;

    DisplayModeConfig config = DisplayModeConfig.builder()
        .keyValuePairComment(CommentVisibility.COMMENT)
        .build();

    YamlWorkbookWriter writer = YamlWorkbookWriter.builder()
        .outputMode(OutputMode.DISPLAY_MODE)
        .displayModeConfig(config)
        .build();

    Workbook workbook = writer.toWorkbook(new java.io.StringReader(yaml));
    Sheet sheet = workbook.getSheetAt(0);

    writeExcelFile(workbook, "display_mode_all_comment.xlsx");

    int rowNum = 0;

    // Row 0: ---
    assertEquals("---", sheet.getRow(rowNum++).getCell(0).getStringCellValue());

    // Row 1: # First comment (KEY_VALUE_PAIR comment shown)
    Row row1 = sheet.getRow(rowNum++);
    assertTrue(row1.getCell(0).getStringCellValue().startsWith("#"));

    // Row 2: # Second comment (KEY_VALUE_PAIR comment shown)
    Row row2 = sheet.getRow(rowNum++);
    assertTrue(row2.getCell(0).getStringCellValue().startsWith("#"));

    // Row 3: name | John
    Row row3 = sheet.getRow(rowNum++);
    assertEquals("name", row3.getCell(0).getStringCellValue());
    assertEquals("John", row3.getCell(1).getStringCellValue());

    workbook.close();
  }

  @Test
  void testYamlOrientedModeUnchanged() throws IOException {
    // Verify YAML_ORIENTED mode still works as before (all comments shown)
    String yaml = """
        # Block comment
        name: John # inline comment
        """;

    YamlWorkbookWriter writer = YamlWorkbookWriter.builder()
        .outputMode(OutputMode.YAML_ORIENTED)
        .build();

    Workbook workbook = writer.toWorkbook(new java.io.StringReader(yaml));
    Sheet sheet = workbook.getSheetAt(0);

    writeExcelFile(workbook, "yaml_oriented_mode.xlsx");

    int rowNum = 0;

    // Row 0: --- (in YAML_ORIENTED, document comments go before frontmatter)
    // But block comments before first key go to key node, not root node
    assertEquals("---", sheet.getRow(rowNum++).getCell(0).getStringCellValue());

    // Row 1: # Block comment (block comment on key node)
    Row row1 = sheet.getRow(rowNum++);
    assertTrue(row1.getCell(0).getStringCellValue().startsWith("#"));

    // Row 2: name | John | # inline comment
    Row row2 = sheet.getRow(rowNum++);
    assertEquals("name", row2.getCell(0).getStringCellValue());
    assertEquals("John", row2.getCell(1).getStringCellValue());
    assertTrue(row2.getCell(2).getStringCellValue().startsWith("#"));

    workbook.close();
  }

  // ==================== Writer Reuse Tests ====================

  @Test
  void testWriterReuseWithYamlContent() throws IOException {
    // Verify that calling toWorkbook() multiple times on the same writer instance
    // produces independent results without state accumulation
    String yaml1 = """
        name: First
        value: 1
        """;
    String yaml2 = """
        name: Second
        value: 2
        extra: field
        """;

    YamlWorkbookWriter writer = YamlWorkbookWriter.builder()
        .outputMode(OutputMode.YAML_ORIENTED)
        .build();

    // First call
    Workbook workbook1 = writer.toWorkbook(new java.io.StringReader(yaml1));
    assertNotNull(workbook1);
    assertEquals(1, workbook1.getNumberOfSheets());
    Sheet sheet1 = workbook1.getSheetAt(0);
    assertEquals(2, sheet1.getLastRowNum()); // frontmatter + 2 rows

    // Second call on same writer instance
    Workbook workbook2 = writer.toWorkbook(new java.io.StringReader(yaml2));
    assertNotNull(workbook2);
    assertEquals(1, workbook2.getNumberOfSheets()); // Should be 1, not 2
    Sheet sheet2 = workbook2.getSheetAt(0);
    assertEquals(3, sheet2.getLastRowNum()); // frontmatter + 3 rows

    // Verify content is independent
    assertEquals("name", sheet1.getRow(1).getCell(0).getStringCellValue());
    assertEquals("First", sheet1.getRow(1).getCell(1).getStringCellValue());

    assertEquals("name", sheet2.getRow(1).getCell(0).getStringCellValue());
    assertEquals("Second", sheet2.getRow(1).getCell(1).getStringCellValue());

    workbook1.close();
    workbook2.close();
  }

  @Test
  void testWriterReuseStateIsReset() throws IOException {
    // Test that internal visibleSheets list is reset between calls
    // by verifying each workbook has exactly one sheet
    String yaml = """
        name: Test
        value: 123
        """;

    YamlWorkbookWriter writer = YamlWorkbookWriter.builder()
        .outputMode(OutputMode.YAML_ORIENTED)
        .build();

    // Call toWorkbook multiple times
    for (int i = 0; i < 3; i++) {
      Workbook workbook = writer.toWorkbook(new java.io.StringReader(yaml));
      // Each workbook should have exactly 1 sheet, not accumulating
      assertEquals(1, workbook.getNumberOfSheets(),
          "Workbook " + i + " should have exactly 1 sheet");
      workbook.close();
    }
  }

  @Test
  void testWriterReuseDataCollectMode() throws IOException {
    // Test reuse in DATA_COLLECT mode
    String jsonSchema = """
        {
          "type": "object",
          "properties": {
            "name": { "type": "string" },
            "age": { "type": "integer" }
          }
        }
        """;

    YamlWorkbookWriter writer = YamlWorkbookWriter.builder()
        .outputMode(OutputMode.FORM_MODE)
        .jsonSchema(jsonSchema)
        .build();

    // First call
    Workbook workbook1 = writer.toWorkbook();
    assertNotNull(workbook1);
    assertEquals(1, workbook1.getNumberOfSheets());
    int rowCount1 = workbook1.getSheetAt(0).getLastRowNum();

    // Second call on same writer instance
    Workbook workbook2 = writer.toWorkbook();
    assertNotNull(workbook2);
    assertEquals(1, workbook2.getNumberOfSheets()); // Should be 1, not 2
    int rowCount2 = workbook2.getSheetAt(0).getLastRowNum();

    // Row counts should be identical
    assertEquals(rowCount1, rowCount2);

    workbook1.close();
    workbook2.close();
  }

  private String loadYaml(String resourcePath) throws IOException {
    try (InputStream is = getClass().getClassLoader().getResourceAsStream(resourcePath)) {
      if (is == null) {
        throw new IOException("Resource not found: " + resourcePath);
      }
      return new String(is.readAllBytes(), StandardCharsets.UTF_8);
    }
  }

  private void writeExcelFile(Workbook workbook, String filename) throws IOException {
    if (WRITE_EXCEL_FILES) {
      Path filePath = OUTPUT_DIR.resolve(filename);
      try (FileOutputStream fos = new FileOutputStream(filePath.toFile())) {
        workbook.write(fos);
      }
    }
  }

}
