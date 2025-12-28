package com.github.wnameless.workbook.yamlworkbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.jupiter.api.BeforeAll;
import org.junit.jupiter.api.Test;

/**
 * Test that writes YAML files to Excel workbooks in target/test-excel for visual inspection.
 */
class DisplayModeExcelOutputTest {

  private static final Path OUTPUT_DIR = Paths.get("target/test-excel");

  @BeforeAll
  static void setup() throws IOException {
    Files.createDirectories(OUTPUT_DIR);
  }

  @Test
  void writeCommentsYamlToExcelWithDisplayMode() throws IOException {
    InputStream is = getClass().getResourceAsStream("/yaml/comments-output.yaml");

    // Test with default DISPLAY_MODE config
    YamlWorkbookWriter writer =
        YamlWorkbookWriter.builder().outputMode(OutputMode.DISPLAY_MODE).build();

    Workbook workbook = writer.toWorkbook(new InputStreamReader(is, StandardCharsets.UTF_8));

    Path outputPath = OUTPUT_DIR.resolve("comments-display-mode.xlsx");
    try (FileOutputStream fos = new FileOutputStream(outputPath.toFile())) {
      workbook.write(fos);
    }
    workbook.close();
    is.close();

    System.out.println("Written to: " + outputPath.toAbsolutePath());
  }

  @Test
  void writeCommentsYamlToExcelWithDisplayModeDisplayName() throws IOException {
    InputStream is = getClass().getResourceAsStream("/yaml/comments-output.yaml");

    // Test with DISPLAY_NAME for all comment types
    DisplayModeConfig config =
        DisplayModeConfig.builder().keyComment(CommentDisplayOption.DISPLAY_NAME)
            .valueComment(CommentDisplayOption.DISPLAY_NAME)
            .mappingComment(CommentDisplayOption.DISPLAY_NAME)
            .sequenceComment(CommentDisplayOption.DISPLAY_NAME).build();

    YamlWorkbookWriter writer = YamlWorkbookWriter.builder().outputMode(OutputMode.DISPLAY_MODE)
        .displayModeConfig(config).build();

    Workbook workbook = writer.toWorkbook(new InputStreamReader(is, StandardCharsets.UTF_8));

    Path outputPath = OUTPUT_DIR.resolve("comments-display-mode-display-name.xlsx");
    try (FileOutputStream fos = new FileOutputStream(outputPath.toFile())) {
      workbook.write(fos);
    }
    workbook.close();
    is.close();

    System.out.println("Written to: " + outputPath.toAbsolutePath());
  }

  @Test
  void writeCommentsYamlToExcelWithDisplayModeHidden() throws IOException {
    InputStream is = getClass().getResourceAsStream("/yaml/comments-output.yaml");

    // Test with HIDDEN for all comment types
    DisplayModeConfig config = DisplayModeConfig.builder().keyComment(CommentDisplayOption.HIDDEN)
        .valueComment(CommentDisplayOption.HIDDEN).mappingComment(CommentDisplayOption.HIDDEN)
        .sequenceComment(CommentDisplayOption.HIDDEN).build();

    YamlWorkbookWriter writer = YamlWorkbookWriter.builder().outputMode(OutputMode.DISPLAY_MODE)
        .displayModeConfig(config).build();

    Workbook workbook = writer.toWorkbook(new InputStreamReader(is, StandardCharsets.UTF_8));

    Path outputPath = OUTPUT_DIR.resolve("comments-display-mode-hidden.xlsx");
    try (FileOutputStream fos = new FileOutputStream(outputPath.toFile())) {
      workbook.write(fos);
    }
    workbook.close();
    is.close();

    System.out.println("Written to: " + outputPath.toAbsolutePath());
  }

  @Test
  void writeCommentsYamlToExcelWithDisplayModeComment() throws IOException {
    InputStream is = getClass().getResourceAsStream("/yaml/comments-output.yaml");

    // Test with COMMENT for all comment types
    DisplayModeConfig config = DisplayModeConfig.builder().keyComment(CommentDisplayOption.COMMENT)
        .valueComment(CommentDisplayOption.COMMENT).mappingComment(CommentDisplayOption.COMMENT)
        .sequenceComment(CommentDisplayOption.COMMENT).build();

    YamlWorkbookWriter writer = YamlWorkbookWriter.builder().outputMode(OutputMode.DISPLAY_MODE)
        .displayModeConfig(config).build();

    Workbook workbook = writer.toWorkbook(new InputStreamReader(is, StandardCharsets.UTF_8));

    Path outputPath = OUTPUT_DIR.resolve("comments-display-mode-comment.xlsx");
    try (FileOutputStream fos = new FileOutputStream(outputPath.toFile())) {
      workbook.write(fos);
    }
    workbook.close();
    is.close();

    System.out.println("Written to: " + outputPath.toAbsolutePath());
  }

  @Test
  void writeCommentsYamlToExcelWithYamlOriented() throws IOException {
    InputStream is = getClass().getResourceAsStream("/yaml/comments-output.yaml");

    // Test with YAML_ORIENTED mode for comparison
    YamlWorkbookWriter writer =
        YamlWorkbookWriter.builder().outputMode(OutputMode.YAML_ORIENTED).build();

    Workbook workbook = writer.toWorkbook(new InputStreamReader(is, StandardCharsets.UTF_8));

    Path outputPath = OUTPUT_DIR.resolve("comments-yaml-oriented.xlsx");
    try (FileOutputStream fos = new FileOutputStream(outputPath.toFile())) {
      workbook.write(fos);
    }
    workbook.close();
    is.close();

    System.out.println("Written to: " + outputPath.toAbsolutePath());
  }

}
