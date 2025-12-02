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
