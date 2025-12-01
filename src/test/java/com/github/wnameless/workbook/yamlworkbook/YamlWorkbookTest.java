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

class YamlWorkbookTest {

  private static final boolean WRITE_EXCEL_FILES = true;
  private static final Path OUTPUT_DIR = Paths.get("target/test-excel");

  private YamlWorkbook yamlWorkbook;

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

    // --- (frontmatter)
    Row row0 = sheet.getRow(0);
    assertEquals("---", row0.getCell(0).getStringCellValue());

    // # This is a configuration file
    Row row1 = sheet.getRow(1);
    assertTrue(row1.getCell(0).getStringCellValue().startsWith("#"));

    // database:
    Row row2 = sheet.getRow(2);
    assertEquals("database", row2.getCell(0).getStringCellValue());

    // # Database connection settings
    Row row3 = sheet.getRow(3);
    assertTrue(row3.getCell(1).getStringCellValue().startsWith("#"));

    // host: localhost # primary host
    Row row4 = sheet.getRow(4);
    assertEquals("host", row4.getCell(1).getStringCellValue());
    assertEquals("localhost", row4.getCell(2).getStringCellValue());
    assertEquals("# primary host", row4.getCell(3).getStringCellValue());

    // port: 5432 # default PostgreSQL port
    Row row5 = sheet.getRow(5);
    assertEquals("port", row5.getCell(1).getStringCellValue());
    assertEquals("5432", row5.getCell(2).getStringCellValue());
    assertEquals("# default PostgreSQL port", row5.getCell(3).getStringCellValue());

    // # Credentials
    Row row6 = sheet.getRow(6);
    assertTrue(row6.getCell(1).getStringCellValue().startsWith("#"));

    // username: admin (no inline comment)
    Row row7 = sheet.getRow(7);
    assertEquals("username", row7.getCell(1).getStringCellValue());
    assertEquals("admin", row7.getCell(2).getStringCellValue());
    assertNull(row7.getCell(3));

    // password: secret # change in production
    Row row8 = sheet.getRow(8);
    assertEquals("password", row8.getCell(1).getStringCellValue());
    assertEquals("secret", row8.getCell(2).getStringCellValue());
    assertEquals("# change in production", row8.getCell(3).getStringCellValue());

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
