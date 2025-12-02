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
  private WorkbookPrintMode workbookPrintMode = WorkbookPrintMode.WORKBOOK_PRONE;
  @Builder.Default
  private WorkbookSymbol workbookSymbol = WorkbookSymbol.DEFAULT;
  @Builder.Default
  private YamlNodeSheetClassifier yamlNodeSheetClassifier = YamlNodeSheetClassifier.DEFAULT;
  @Builder.Default
  private WorkbookSheetNameStrategy workbookSheetNameStrategy = WorkbookSheetNameStrategy.DEFAULT;

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
      workbook.createSheet(workbookSheetNameStrategy.apply(0));
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
    var sheetIdx = yamlNodeSheetClassifier.apply(node, nodeIdx);
    while (workbook.getNumberOfSheets() <= sheetIdx) {
      workbook.createSheet(workbookSheetNameStrategy.apply(workbook.getNumberOfSheets()));
    }
    var sheet = workbook.getSheetAt(sheetIdx);
    writeFrontmatter(sheet);
    traverseAndPrintNode(node, sheet, 0);
    nodeIdx++;
  }

  private void writeFrontmatter(Sheet sheet) {
    Row row = sheet.createRow(sheet.getLastRowNum() + 1);
    Cell cell = row.createCell(0);
    cell.setCellValue(workbookSymbol.getFrontmatter());
  }

  private void traverseAndPrintNode(Node node, Sheet sheet, int indentLevel) {
    if (node == null) {
      return;
    }

    writeComments(node.getBlockComments(), sheet, indentLevel);

    if (node instanceof ScalarNode scalarNode) {
      traverseScalarNode(scalarNode, sheet, indentLevel);
    } else if (node instanceof MappingNode mappingNode) {
      traverseMappingNode(mappingNode, sheet, indentLevel);
    } else if (node instanceof SequenceNode sequenceNode) {
      traverseSequenceNode(sequenceNode, sheet, indentLevel);
    }

    writeComments(node.getEndComments(), sheet, indentLevel);
  }

  private void traverseScalarNode(ScalarNode node, Sheet sheet, int indentLevel) {
    Row row = sheet.createRow(sheet.getLastRowNum() + 1);
    int cellIndex = indentLevel * workbookSymbol.getIndentationCellNum();
    Cell cell = row.createCell(cellIndex);
    cell.setCellValue(escapeValueIfNeeded(node.getValue()));
  }

  private void traverseMappingNode(MappingNode node, Sheet sheet, int indentLevel) {
    for (NodeTuple tuple : node.getValue()) {
      Node keyNode = tuple.getKeyNode();
      Node valueNode = tuple.getValueNode();

      writeComments(keyNode.getBlockComments(), sheet, indentLevel);

      if (keyNode instanceof ScalarNode scalarKey) {
        Row row = sheet.createRow(sheet.getLastRowNum() + 1);
        int cellIndex = indentLevel * workbookSymbol.getIndentationCellNum();
        Cell keyCell = row.createCell(cellIndex);
        keyCell.setCellValue(scalarKey.getValue());

        if (valueNode instanceof ScalarNode scalarValue) {
          int nextCellIndex = cellIndex + 1;
          nextCellIndex = writeInlineComments(keyNode.getInLineComments(), row, nextCellIndex);
          Cell valueCell = row.createCell(nextCellIndex);
          valueCell.setCellValue(escapeValueIfNeeded(scalarValue.getValue()));
          writeInlineComments(valueNode.getInLineComments(), row, nextCellIndex + 1);
        } else {
          writeInlineComments(keyNode.getInLineComments(), row, cellIndex + 1);
          traverseAndPrintNode(valueNode, sheet, indentLevel + 1);
        }
      } else {
        traverseAndPrintNode(keyNode, sheet, indentLevel);
        traverseAndPrintNode(valueNode, sheet, indentLevel + 1);
      }

      writeComments(keyNode.getEndComments(), sheet, indentLevel);
    }
  }

  private void traverseSequenceNode(SequenceNode node, Sheet sheet, int indentLevel) {
    for (Node item : node.getValue()) {
      writeComments(item.getBlockComments(), sheet, indentLevel);

      Row row = sheet.createRow(sheet.getLastRowNum() + 1);
      int cellIndex = indentLevel * workbookSymbol.getIndentationCellNum();
      Cell itemMarkCell = row.createCell(cellIndex);
      itemMarkCell.setCellValue(workbookSymbol.getItemMark());

      if (item instanceof ScalarNode scalarItem) {
        Cell valueCell = row.createCell(cellIndex + 1);
        valueCell.setCellValue(escapeValueIfNeeded(scalarItem.getValue()));
        writeInlineComments(item.getInLineComments(), row, cellIndex + 2);
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
      int cellIndex = indentLevel * workbookSymbol.getIndentationCellNum();
      Cell cell = row.createCell(cellIndex);
      cell.setCellValue(workbookSymbol.getCommentMark() + " " + comment.getValue().trim());
    }
  }

  private int writeInlineComments(List<CommentLine> comments, Row row, int startCellIndex) {
    if (comments == null || comments.isEmpty()) {
      return startCellIndex;
    }

    int cellIndex = startCellIndex;
    for (CommentLine comment : comments) {
      Cell cell = row.createCell(cellIndex++);
      cell.setCellValue(workbookSymbol.getCommentMark() + " " + comment.getValue().trim());
    }
    return cellIndex;
  }

  private String escapeValueIfNeeded(String value) {
    if (value == null) {
      return null;
    }
    // Only escape if value STARTS with comment mark or escape mark
    if (value.startsWith(workbookSymbol.getCommentMark())
        || value.startsWith(workbookSymbol.getValueEscapeMark())) {
      return workbookSymbol.getValueEscapeMark() + value;
    }
    return value;
  }

}
