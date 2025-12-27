package com.github.wnameless.workbook.yamlworkbook;

import static org.junit.jupiter.api.Assertions.*;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.StringReader;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.List;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.jupiter.api.BeforeAll;
import org.junit.jupiter.api.Test;
import org.yaml.snakeyaml.nodes.MappingNode;
import org.yaml.snakeyaml.nodes.Node;
import org.yaml.snakeyaml.nodes.ScalarNode;
import org.yaml.snakeyaml.nodes.SequenceNode;

class PrefixIndentationTest {

  private static final boolean WRITE_EXCEL_FILES = true;
  private static final Path OUTPUT_DIR = Paths.get("target/test-excel");

  @BeforeAll
  static void setUpOnce() throws IOException {
    if (WRITE_EXCEL_FILES) {
      Files.createDirectories(OUTPUT_DIR);
    }
  }

  // ==================== DefaultIndentPrefixStrategy Tests ====================

  @Test
  void testDefaultPrefixStrategyGenerate() {
    IndentPrefixStrategy strategy = IndentPrefixStrategy.DEFAULT;

    assertEquals("", strategy.generatePrefix(0));
    assertEquals("1>", strategy.generatePrefix(1));
    assertEquals("2>", strategy.generatePrefix(2));
    assertEquals("3>", strategy.generatePrefix(3));
    assertEquals("10>", strategy.generatePrefix(10));
  }

  @Test
  void testDefaultPrefixStrategyParse() {
    IndentPrefixStrategy strategy = IndentPrefixStrategy.DEFAULT;

    assertEquals(0, strategy.parsePrefix(""));
    assertEquals(0, strategy.parsePrefix(null));
    assertEquals(1, strategy.parsePrefix("1>"));
    assertEquals(2, strategy.parsePrefix("2>"));
    assertEquals(3, strategy.parsePrefix("3>"));
    assertEquals(10, strategy.parsePrefix("10>"));

    // Invalid prefixes
    assertEquals(-1, strategy.parsePrefix("1"));
    assertEquals(-1, strategy.parsePrefix(">"));
    assertEquals(-1, strategy.parsePrefix("abc"));
    assertEquals(-1, strategy.parsePrefix("0>"));
    assertEquals(-1, strategy.parsePrefix("-1>"));
  }

  @Test
  void testDefaultPrefixStrategyIsPrefix() {
    IndentPrefixStrategy strategy = IndentPrefixStrategy.DEFAULT;

    assertTrue(strategy.isPrefix(""));
    assertTrue(strategy.isPrefix("1>"));
    assertTrue(strategy.isPrefix("2>"));

    assertFalse(strategy.isPrefix("abc"));
    assertFalse(strategy.isPrefix("1"));
  }

  // ==================== Writer PREFIX Mode Tests ====================

  @Test
  void testPrefixModeSimpleMapping() throws IOException {
    String yaml = """
        name: John
        age: 30
        """;

    YamlWorkbookWriter writer = YamlWorkbook.prefixWriterBuilder().build();
    Workbook workbook = writer.toWorkbook(new StringReader(yaml));

    Sheet sheet = workbook.getSheetAt(0);

    // Row 0: --- (frontmatter, no prefix)
    assertEquals("---", sheet.getRow(0).getCell(0).getStringCellValue());

    // Row 1: name | John (level 0, no prefix)
    Row row1 = sheet.getRow(1);
    assertEquals("name", row1.getCell(0).getStringCellValue());
    assertEquals("John", row1.getCell(1).getStringCellValue());

    // Row 2: age | 30 (level 0, no prefix)
    Row row2 = sheet.getRow(2);
    assertEquals("age", row2.getCell(0).getStringCellValue());
    assertEquals("30", row2.getCell(1).getStringCellValue());

    writeExcelFile(workbook, "prefix_simple.xlsx");
    workbook.close();
  }

  @Test
  void testPrefixModeNestedMapping() throws IOException {
    String yaml = """
        person:
          name: John Doe
          address:
            city: New York
            zip: 10001
        """;

    YamlWorkbookWriter writer = YamlWorkbook.prefixWriterBuilder().build();
    Workbook workbook = writer.toWorkbook(new StringReader(yaml));

    Sheet sheet = workbook.getSheetAt(0);

    // Row 0: --- (frontmatter)
    assertEquals("---", sheet.getRow(0).getCell(0).getStringCellValue());

    // Row 1: person (level 0, no prefix)
    Row row1 = sheet.getRow(1);
    assertEquals("person", row1.getCell(0).getStringCellValue());

    // Row 2: 1> | name | John Doe (level 1)
    Row row2 = sheet.getRow(2);
    assertEquals("1>", row2.getCell(0).getStringCellValue());
    assertEquals("name", row2.getCell(1).getStringCellValue());
    assertEquals("John Doe", row2.getCell(2).getStringCellValue());

    // Row 3: 1> | address (level 1)
    Row row3 = sheet.getRow(3);
    assertEquals("1>", row3.getCell(0).getStringCellValue());
    assertEquals("address", row3.getCell(1).getStringCellValue());

    // Row 4: 2> | city | New York (level 2)
    Row row4 = sheet.getRow(4);
    assertEquals("2>", row4.getCell(0).getStringCellValue());
    assertEquals("city", row4.getCell(1).getStringCellValue());
    assertEquals("New York", row4.getCell(2).getStringCellValue());

    // Row 5: 2> | zip | 10001 (level 2)
    Row row5 = sheet.getRow(5);
    assertEquals("2>", row5.getCell(0).getStringCellValue());
    assertEquals("zip", row5.getCell(1).getStringCellValue());
    assertEquals("10001", row5.getCell(2).getStringCellValue());

    writeExcelFile(workbook, "prefix_nested.xlsx");
    workbook.close();
  }

  @Test
  void testPrefixModeSequence() throws IOException {
    String yaml = """
        fruits:
          - apple
          - banana
          - cherry
        """;

    YamlWorkbookWriter writer = YamlWorkbook.prefixWriterBuilder().build();
    Workbook workbook = writer.toWorkbook(new StringReader(yaml));

    Sheet sheet = workbook.getSheetAt(0);

    // Row 0: ---
    assertEquals("---", sheet.getRow(0).getCell(0).getStringCellValue());

    // Row 1: fruits (level 0)
    assertEquals("fruits", sheet.getRow(1).getCell(0).getStringCellValue());

    // Row 2: 1> | - | apple (level 1)
    Row row2 = sheet.getRow(2);
    assertEquals("1>", row2.getCell(0).getStringCellValue());
    assertEquals("-", row2.getCell(1).getStringCellValue());
    assertEquals("apple", row2.getCell(2).getStringCellValue());

    // Row 3: 1> | - | banana
    Row row3 = sheet.getRow(3);
    assertEquals("1>", row3.getCell(0).getStringCellValue());
    assertEquals("-", row3.getCell(1).getStringCellValue());
    assertEquals("banana", row3.getCell(2).getStringCellValue());

    // Row 4: 1> | - | cherry
    Row row4 = sheet.getRow(4);
    assertEquals("1>", row4.getCell(0).getStringCellValue());
    assertEquals("-", row4.getCell(1).getStringCellValue());
    assertEquals("cherry", row4.getCell(2).getStringCellValue());

    writeExcelFile(workbook, "prefix_sequence.xlsx");
    workbook.close();
  }

  @Test
  void testPrefixModeDeeplyNested() throws IOException {
    String yaml = """
        level0:
          level1:
            level2:
              level3:
                level4: deep value
        """;

    YamlWorkbookWriter writer = YamlWorkbook.prefixWriterBuilder().build();
    Workbook workbook = writer.toWorkbook(new StringReader(yaml));

    Sheet sheet = workbook.getSheetAt(0);

    // Row 0: ---
    assertEquals("---", sheet.getRow(0).getCell(0).getStringCellValue());

    // Row 1: level0 (level 0, no prefix)
    assertEquals("level0", sheet.getRow(1).getCell(0).getStringCellValue());

    // Row 2: 1> | level1
    assertEquals("1>", sheet.getRow(2).getCell(0).getStringCellValue());
    assertEquals("level1", sheet.getRow(2).getCell(1).getStringCellValue());

    // Row 3: 2> | level2
    assertEquals("2>", sheet.getRow(3).getCell(0).getStringCellValue());
    assertEquals("level2", sheet.getRow(3).getCell(1).getStringCellValue());

    // Row 4: 3> | level3
    assertEquals("3>", sheet.getRow(4).getCell(0).getStringCellValue());
    assertEquals("level3", sheet.getRow(4).getCell(1).getStringCellValue());

    // Row 5: 4> | level4 | deep value
    assertEquals("4>", sheet.getRow(5).getCell(0).getStringCellValue());
    assertEquals("level4", sheet.getRow(5).getCell(1).getStringCellValue());
    assertEquals("deep value", sheet.getRow(5).getCell(2).getStringCellValue());

    writeExcelFile(workbook, "prefix_deep.xlsx");
    workbook.close();
  }

  // ==================== Roundtrip Tests ====================

  @Test
  void testPrefixModeRoundtripSimple() throws IOException {
    String yaml = """
        name: John
        age: 30
        city: New York
        """;

    // Write with prefix mode
    YamlWorkbookWriter writer = YamlWorkbook.prefixWriterBuilder().build();
    Workbook workbook = writer.toWorkbook(new StringReader(yaml));

    // Read with prefix mode
    YamlWorkbookReader reader = YamlWorkbook.prefixReaderBuilder().build();
    List<Node> nodes = reader.fromWorkbook(workbook);

    assertEquals(1, nodes.size());
    assertTrue(nodes.get(0) instanceof MappingNode);
    MappingNode mapping = (MappingNode) nodes.get(0);
    assertEquals(3, mapping.getValue().size());

    workbook.close();
  }

  @Test
  void testPrefixModeRoundtripNested() throws IOException {
    String yaml = """
        person:
          name: Jane
          address:
            street: 123 Main St
            city: Boston
        """;

    // Write with prefix mode
    YamlWorkbookWriter writer = YamlWorkbook.prefixWriterBuilder().build();
    Workbook workbook = writer.toWorkbook(new StringReader(yaml));

    writeExcelFile(workbook, "prefix_roundtrip_nested.xlsx");

    // Read with prefix mode
    YamlWorkbookReader reader = YamlWorkbook.prefixReaderBuilder().build();
    List<Node> nodes = reader.fromWorkbook(workbook);

    assertEquals(1, nodes.size());
    assertTrue(nodes.get(0) instanceof MappingNode);

    MappingNode root = (MappingNode) nodes.get(0);
    assertEquals(1, root.getValue().size());

    // person
    ScalarNode personKey = (ScalarNode) root.getValue().get(0).getKeyNode();
    assertEquals("person", personKey.getValue());

    MappingNode personValue = (MappingNode) root.getValue().get(0).getValueNode();
    assertEquals(2, personValue.getValue().size());

    // name
    ScalarNode nameKey = (ScalarNode) personValue.getValue().get(0).getKeyNode();
    assertEquals("name", nameKey.getValue());
    ScalarNode nameValue = (ScalarNode) personValue.getValue().get(0).getValueNode();
    assertEquals("Jane", nameValue.getValue());

    // address
    ScalarNode addressKey = (ScalarNode) personValue.getValue().get(1).getKeyNode();
    assertEquals("address", addressKey.getValue());
    MappingNode addressValue = (MappingNode) personValue.getValue().get(1).getValueNode();
    assertEquals(2, addressValue.getValue().size());

    workbook.close();
  }

  @Test
  void testPrefixModeRoundtripSequence() throws IOException {
    String yaml = """
        items:
          - first
          - second
          - third
        """;

    // Write with prefix mode
    YamlWorkbookWriter writer = YamlWorkbook.prefixWriterBuilder().build();
    Workbook workbook = writer.toWorkbook(new StringReader(yaml));

    // Read with prefix mode
    YamlWorkbookReader reader = YamlWorkbook.prefixReaderBuilder().build();
    List<Node> nodes = reader.fromWorkbook(workbook);

    assertEquals(1, nodes.size());
    MappingNode root = (MappingNode) nodes.get(0);
    assertEquals(1, root.getValue().size());

    SequenceNode items = (SequenceNode) root.getValue().get(0).getValueNode();
    assertEquals(3, items.getValue().size());

    assertEquals("first", ((ScalarNode) items.getValue().get(0)).getValue());
    assertEquals("second", ((ScalarNode) items.getValue().get(1)).getValue());
    assertEquals("third", ((ScalarNode) items.getValue().get(2)).getValue());

    workbook.close();
  }

  // ==================== Custom Prefix Strategy Test ====================

  @Test
  void testCustomPrefixStrategy() throws IOException {
    // Custom strategy using "L1:", "L2:", etc.
    IndentPrefixStrategy customStrategy = new IndentPrefixStrategy() {
      @Override
      public String generatePrefix(int indentLevel) {
        return indentLevel <= 0 ? "" : "L" + indentLevel + ":";
      }

      @Override
      public int parsePrefix(String prefix) {
        if (prefix == null || prefix.isEmpty()) return 0;
        if (!prefix.startsWith("L") || !prefix.endsWith(":")) return -1;
        try {
          int level = Integer.parseInt(prefix.substring(1, prefix.length() - 1));
          return level > 0 ? level : -1;
        } catch (NumberFormatException e) {
          return -1;
        }
      }
    };

    String yaml = """
        root:
          nested: value
        """;

    YamlWorkbookWriter writer = YamlWorkbookWriter.builder()
        .indentationMode(IndentationMode.PREFIX)
        .indentPrefixStrategy(customStrategy)
        .build();
    Workbook workbook = writer.toWorkbook(new StringReader(yaml));

    Sheet sheet = workbook.getSheetAt(0);

    // Row 1: root (level 0, no prefix)
    assertEquals("root", sheet.getRow(1).getCell(0).getStringCellValue());

    // Row 2: L1: | nested | value (level 1 with custom prefix)
    assertEquals("L1:", sheet.getRow(2).getCell(0).getStringCellValue());
    assertEquals("nested", sheet.getRow(2).getCell(1).getStringCellValue());
    assertEquals("value", sheet.getRow(2).getCell(2).getStringCellValue());

    // Roundtrip with custom strategy
    YamlWorkbookReader reader = YamlWorkbookReader.builder()
        .indentationMode(IndentationMode.PREFIX)
        .indentPrefixStrategy(customStrategy)
        .build();
    List<Node> nodes = reader.fromWorkbook(workbook);

    assertEquals(1, nodes.size());
    MappingNode root = (MappingNode) nodes.get(0);
    assertEquals(1, root.getValue().size());
    assertEquals("root", ((ScalarNode) root.getValue().get(0).getKeyNode()).getValue());

    writeExcelFile(workbook, "prefix_custom.xlsx");
    workbook.close();
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
