package com.github.wnameless.workbook.yamlworkbook;

import java.io.StringReader;
import java.util.ArrayList;
import java.util.List;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.yaml.snakeyaml.LoaderOptions;
import org.yaml.snakeyaml.Yaml;
import org.yaml.snakeyaml.comments.CommentLine;
import org.yaml.snakeyaml.nodes.MappingNode;
import org.yaml.snakeyaml.nodes.Node;
import org.yaml.snakeyaml.nodes.NodeTuple;
import org.yaml.snakeyaml.nodes.ScalarNode;
import org.yaml.snakeyaml.nodes.SequenceNode;
import lombok.Builder;

@Builder
public class YamlWorkbookWriter {

  @Builder.Default
  private PrintMode printMode = PrintMode.YAML_ORIENTED;
  @Builder.Default
  private DisplayModeConfig displayModeConfig = DisplayModeConfig.DEFAULT;
  @Builder.Default
  private WorkbookSyntax workbookSyntax = WorkbookSyntax.DEFAULT;
  @Builder.Default
  private NodeToSheetMapper nodeToSheetMapper = NodeToSheetMapper.DEFAULT;
  @Builder.Default
  private SheetNameStrategy sheetNameStrategy = SheetNameStrategy.DEFAULT;

  public Workbook toWorkbook(StringReader yamlContent, StringReader... yamlContents) {
    var workbook = new XSSFWorkbook();

    LoaderOptions options = new LoaderOptions();
    options.setProcessComments(true);
    Yaml yaml = new Yaml(options);

    List<Iterable<Node>> nodeIters = new ArrayList<>();
    nodeIters.add(yaml.composeAll(yamlContent));
    for (StringReader content : yamlContents) {
      nodeIters.add(yaml.composeAll(content));
    }

    processNodes(nodeIters, workbook);

    if (workbook.getNumberOfSheets() == 0) {
      workbook.createSheet(sheetNameStrategy.apply(0));
    }
    return workbook;
  }

  private void processNodes(List<Iterable<Node>> nodeIters, Workbook workbook) {
    int nodeIdx = 0;
    for (var nodeIter : nodeIters) {
      for (var node : nodeIter) {
        processNode(node, workbook, nodeIdx);
      }
    }
  }

  private void processNode(Node node, Workbook workbook, int nodeIdx) {
    var sheetIdx = nodeToSheetMapper.apply(node, nodeIdx);
    while (workbook.getNumberOfSheets() <= sheetIdx) {
      workbook.createSheet(sheetNameStrategy.apply(workbook.getNumberOfSheets()));
    }
    var sheet = workbook.getSheetAt(sheetIdx);

    // Handle document-level comments (before frontmatter)
    if (isDisplayMode()) {
      if (displayModeConfig.getDocumentComment() == CommentVisibility.COMMENT) {
        writeDocumentComments(node, sheet);
      }
    } else {
      writeDocumentComments(node, sheet);
    }

    writeFrontmatter(sheet);
    traverseAndPrintNodeWithoutBlockComments(node, sheet, 0);
    nodeIdx++;
  }

  private void writeDocumentComments(Node node, Sheet sheet) {
    // Document comments are the block comments of the root node
    writeComments(node.getBlockComments(), sheet, 0);
  }

  private void writeFrontmatter(Sheet sheet) {
    Row row = sheet.createRow(sheet.getLastRowNum() + 1);
    Cell cell = row.createCell(0);
    cell.setCellValue(workbookSyntax.getFrontmatter());
  }

  private void traverseAndPrintNodeWithoutBlockComments(Node node, Sheet sheet, int indentLevel) {
    // Used for root node where block comments are handled as document comments
    if (node == null) {
      return;
    }

    if (node instanceof ScalarNode scalarNode) {
      traverseScalarNode(scalarNode, sheet, indentLevel);
    } else if (node instanceof MappingNode mappingNode) {
      traverseMappingNode(mappingNode, sheet, indentLevel);
    } else if (node instanceof SequenceNode sequenceNode) {
      traverseSequenceNode(sequenceNode, sheet, indentLevel);
    }

    writeComments(node.getEndComments(), sheet, indentLevel);
  }

  private void traverseAndPrintNode(Node node, Sheet sheet, int indentLevel) {
    if (node == null) {
      return;
    }

    // Handle block comments based on node type (OBJECT/ARRAY/VALUE comments)
    if (isDisplayMode()) {
      if (node instanceof MappingNode) {
        writeBlockCommentsInDisplayModeReplaceable(node.getBlockComments(), sheet, indentLevel,
            displayModeConfig.getObjectComment());
      } else if (node instanceof SequenceNode) {
        writeBlockCommentsInDisplayModeReplaceable(node.getBlockComments(), sheet, indentLevel,
            displayModeConfig.getArrayComment());
      } else {
        writeComments(node.getBlockComments(), sheet, indentLevel);
      }
    } else {
      writeComments(node.getBlockComments(), sheet, indentLevel);
    }

    if (node instanceof ScalarNode scalarNode) {
      traverseScalarNode(scalarNode, sheet, indentLevel);
    } else if (node instanceof MappingNode mappingNode) {
      traverseMappingNode(mappingNode, sheet, indentLevel);
    } else if (node instanceof SequenceNode sequenceNode) {
      traverseSequenceNode(sequenceNode, sheet, indentLevel);
    }

    writeComments(node.getEndComments(), sheet, indentLevel);
  }

  private void writeBlockCommentsInDisplayModeReplaceable(List<CommentLine> comments, Sheet sheet,
      int indentLevel, CommentDisplayOption option) {
    if (comments == null || comments.isEmpty()) {
      return;
    }
    switch (option) {
      case DISPLAY_NAME -> {
        // For OBJECT/ARRAY, DISPLAY_NAME shows the comment as a header row
        Row row = sheet.createRow(sheet.getLastRowNum() + 1);
        int cellIndex = indentLevel * workbookSyntax.getIndentationCellNum();
        Cell cell = row.createCell(cellIndex);
        cell.setCellValue(extractCommentText(comments));
      }
      case HIDDEN -> { /* skip comments entirely */ }
      case COMMENT -> writeComments(comments, sheet, indentLevel);
    }
  }

  private void traverseScalarNode(ScalarNode node, Sheet sheet, int indentLevel) {
    Row row = sheet.createRow(sheet.getLastRowNum() + 1);
    int cellIndex = indentLevel * workbookSyntax.getIndentationCellNum();
    Cell cell = row.createCell(cellIndex);
    cell.setCellValue(escapeValueIfNeeded(node.getValue()));
  }

  private void traverseMappingNode(MappingNode node, Sheet sheet, int indentLevel) {
    for (NodeTuple tuple : node.getValue()) {
      Node keyNode = tuple.getKeyNode();
      Node valueNode = tuple.getValueNode();

      // Handle block comments (KEY_VALUE_PAIR comment type)
      if (isDisplayMode()) {
        writeBlockCommentsInDisplayMode(keyNode.getBlockComments(), sheet, indentLevel,
            displayModeConfig.getKeyValuePairComment());
      } else {
        writeComments(keyNode.getBlockComments(), sheet, indentLevel);
      }

      if (keyNode instanceof ScalarNode scalarKey) {
        Row row = sheet.createRow(sheet.getLastRowNum() + 1);
        int cellIndex = indentLevel * workbookSyntax.getIndentationCellNum();
        Cell keyCell = row.createCell(cellIndex);

        // Handle key display
        String keyDisplayValue = scalarKey.getValue();
        if (isDisplayMode()) {
          String keyCommentText = extractCommentText(scalarKey.getInLineComments());
          if (keyCommentText != null) {
            switch (displayModeConfig.getKeyComment()) {
              case DISPLAY_NAME -> keyDisplayValue = keyCommentText;
              case HIDDEN -> { /* keep original */ }
              case COMMENT -> { /* handled separately */ }
            }
          }
        }
        keyCell.setCellValue(keyDisplayValue);

        if (valueNode instanceof ScalarNode scalarValue) {
          int nextCellIndex = cellIndex + 1;

          // Handle key inline comments
          if (isDisplayMode()) {
            if (displayModeConfig.getKeyComment() == CommentDisplayOption.COMMENT) {
              nextCellIndex = writeInlineComments(keyNode.getInLineComments(), row, nextCellIndex);
            }
          } else {
            nextCellIndex = writeInlineComments(keyNode.getInLineComments(), row, nextCellIndex);
          }

          Cell valueCell = row.createCell(nextCellIndex);

          // Handle value display
          String valueDisplayValue = escapeValueIfNeeded(scalarValue.getValue());
          if (isDisplayMode()) {
            String valueCommentText = extractCommentText(scalarValue.getInLineComments());
            if (valueCommentText != null) {
              switch (displayModeConfig.getValueComment()) {
                case DISPLAY_NAME -> valueDisplayValue = valueCommentText;
                case HIDDEN -> { /* keep original */ }
                case COMMENT -> { /* handled separately */ }
              }
            }
          }
          valueCell.setCellValue(valueDisplayValue);

          // Handle value inline comments
          if (isDisplayMode()) {
            if (displayModeConfig.getValueComment() == CommentDisplayOption.COMMENT) {
              writeInlineComments(valueNode.getInLineComments(), row, nextCellIndex + 1);
            }
          } else {
            writeInlineComments(valueNode.getInLineComments(), row, nextCellIndex + 1);
          }
        } else {
          // Handle key inline comments for nested value
          if (isDisplayMode()) {
            if (displayModeConfig.getKeyComment() == CommentDisplayOption.COMMENT) {
              writeInlineComments(keyNode.getInLineComments(), row, cellIndex + 1);
            }
          } else {
            writeInlineComments(keyNode.getInLineComments(), row, cellIndex + 1);
          }
          traverseAndPrintNode(valueNode, sheet, indentLevel + 1);
        }
      } else {
        traverseAndPrintNode(keyNode, sheet, indentLevel);
        traverseAndPrintNode(valueNode, sheet, indentLevel + 1);
      }

      writeComments(keyNode.getEndComments(), sheet, indentLevel);
    }
  }

  private void writeBlockCommentsInDisplayMode(List<CommentLine> comments, Sheet sheet,
      int indentLevel, CommentVisibility visibility) {
    if (comments == null || comments.isEmpty() || visibility == CommentVisibility.HIDDEN) {
      return;
    }
    // COMMENT visibility - write as regular comments
    writeComments(comments, sheet, indentLevel);
  }

  private void traverseSequenceNode(SequenceNode node, Sheet sheet, int indentLevel) {
    for (Node item : node.getValue()) {
      // Handle item block comments
      if (isDisplayMode()) {
        writeBlockCommentsInDisplayMode(item.getBlockComments(), sheet, indentLevel,
            displayModeConfig.getItemComment());
      } else {
        writeComments(item.getBlockComments(), sheet, indentLevel);
      }

      Row row = sheet.createRow(sheet.getLastRowNum() + 1);
      int cellIndex = indentLevel * workbookSyntax.getIndentationCellNum();
      Cell itemMarkCell = row.createCell(cellIndex);
      itemMarkCell.setCellValue(workbookSyntax.getItemMark());

      if (item instanceof ScalarNode scalarItem) {
        Cell valueCell = row.createCell(cellIndex + 1);
        valueCell.setCellValue(escapeValueIfNeeded(scalarItem.getValue()));

        // Handle item inline comments
        if (isDisplayMode()) {
          if (displayModeConfig.getItemComment() == CommentVisibility.COMMENT) {
            writeInlineComments(item.getInLineComments(), row, cellIndex + 2);
          }
        } else {
          writeInlineComments(item.getInLineComments(), row, cellIndex + 2);
        }
      } else {
        traverseAndPrintNode(item, sheet, indentLevel + 1);
      }

      writeComments(item.getEndComments(), sheet, indentLevel);
    }
  }

  private void writeComments(List<CommentLine> comments, Sheet sheet, int indentLevel) {
    if (comments == null || comments.isEmpty()) {
      return;
    }

    for (CommentLine comment : comments) {
      Row row = sheet.createRow(sheet.getLastRowNum() + 1);
      int cellIndex = indentLevel * workbookSyntax.getIndentationCellNum();
      Cell cell = row.createCell(cellIndex);
      cell.setCellValue(workbookSyntax.getCommentMark() + " " + comment.getValue().trim());
    }
  }

  private int writeInlineComments(List<CommentLine> comments, Row row, int startCellIndex) {
    if (comments == null || comments.isEmpty()) {
      return startCellIndex;
    }

    int cellIndex = startCellIndex;
    for (CommentLine comment : comments) {
      Cell cell = row.createCell(cellIndex++);
      cell.setCellValue(workbookSyntax.getCommentMark() + " " + comment.getValue().trim());
    }
    return cellIndex;
  }

  private String escapeValueIfNeeded(String value) {
    if (value == null) {
      return null;
    }
    // Only escape if value STARTS with comment mark or escape mark
    if (value.startsWith(workbookSyntax.getCommentMark())
        || value.startsWith(workbookSyntax.getValueEscapeMark())) {
      return workbookSyntax.getValueEscapeMark() + value;
    }
    return value;
  }

  private boolean isDisplayMode() {
    return printMode == PrintMode.WORKBOOK_DISPLAY;
  }

  private String extractCommentText(List<CommentLine> comments) {
    if (comments == null || comments.isEmpty()) {
      return null;
    }
    // Use the first comment's text as display name
    return comments.get(0).getValue().trim();
  }

}
