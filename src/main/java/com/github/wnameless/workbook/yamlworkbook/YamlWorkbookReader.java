package com.github.wnameless.workbook.yamlworkbook;

import java.util.ArrayList;
import java.util.List;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.yaml.snakeyaml.DumperOptions.FlowStyle;
import org.yaml.snakeyaml.DumperOptions.ScalarStyle;
import org.yaml.snakeyaml.comments.CommentLine;
import org.yaml.snakeyaml.comments.CommentType;
import org.yaml.snakeyaml.nodes.MappingNode;
import org.yaml.snakeyaml.nodes.Node;
import org.yaml.snakeyaml.nodes.NodeTuple;
import org.yaml.snakeyaml.nodes.ScalarNode;
import org.yaml.snakeyaml.nodes.SequenceNode;
import org.yaml.snakeyaml.nodes.Tag;
import lombok.Builder;

@Builder
public class YamlWorkbookReader {

  @Builder.Default
  private WorkbookPrintMode workbookPrintMode = WorkbookPrintMode.WORKBOOK_PRONE;
  @Builder.Default
  private WorkbookSymbol workbookSymbol = WorkbookSymbol.DEFAULT;
  @Builder.Default
  private WorkbookSheetNameStrategy workbookSheetNameStrategy = WorkbookSheetNameStrategy.DEFAULT;

  public List<Node> fromWorkbook(Workbook workbook) {
    var nodeList = new ArrayList<Node>();
    if (workbook == null) return nodeList;

    for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
      if (workbook.getSheetName(i).equals(workbookSheetNameStrategy.apply(i))) {
        processYamlSheet(workbook.getSheetAt(i)).forEach(nodeList::add);
      }
    }

    return nodeList;
  }

  private Iterable<Node> processYamlSheet(Sheet sheet) {
    List<Node> documents = new ArrayList<>();
    if (sheet == null) return documents;

    List<List<Row>> documentRows = splitByFrontmatter(sheet);
    for (List<Row> docRows : documentRows) {
      if (!docRows.isEmpty()) {
        Node docNode = parseRows(docRows, 0, 0, docRows.size());
        if (docNode != null) {
          documents.add(docNode);
        }
      }
    }
    return documents;
  }

  private List<List<Row>> splitByFrontmatter(Sheet sheet) {
    List<List<Row>> documents = new ArrayList<>();
    List<Row> currentDoc = new ArrayList<>();

    for (int i = 0; i <= sheet.getLastRowNum(); i++) {
      Row row = sheet.getRow(i);
      if (row == null) continue;

      String firstCellValue = getCellValue(row, 0);
      if (workbookSymbol.getFrontmatter().equals(firstCellValue)) {
        if (!currentDoc.isEmpty()) {
          documents.add(currentDoc);
          currentDoc = new ArrayList<>();
        }
      } else {
        currentDoc.add(row);
      }
    }

    if (!currentDoc.isEmpty()) {
      documents.add(currentDoc);
    }

    return documents;
  }

  private Node parseRows(List<Row> rows, int indentLevel, int startIdx, int endIdx) {
    if (startIdx >= endIdx) return null;

    List<CommentLine> pendingComments = new ArrayList<>();
    int cellOffset = indentLevel * workbookSymbol.getIndentationCellNum();

    // Find first non-comment row to determine structure type
    int firstContentIdx = startIdx;
    while (firstContentIdx < endIdx) {
      Row row = rows.get(firstContentIdx);
      String firstValue = getCellValue(row, cellOffset);
      if (firstValue != null && !isComment(firstValue)) {
        break;
      }
      if (firstValue != null && isComment(firstValue)) {
        pendingComments.add(createCommentLine(firstValue));
      }
      firstContentIdx++;
    }

    if (firstContentIdx >= endIdx) {
      return null; // Only comments, no content
    }

    Row firstRow = rows.get(firstContentIdx);
    String firstValue = getCellValue(firstRow, cellOffset);

    // Determine if this is a sequence, mapping, or scalar
    if (isItemMark(firstValue)) {
      return parseSequence(rows, indentLevel, startIdx, endIdx, pendingComments);
    } else {
      String secondValue = getCellValue(firstRow, cellOffset + 1);
      if (secondValue != null || hasNestedContent(rows, indentLevel, firstContentIdx, endIdx)) {
        return parseMapping(rows, indentLevel, startIdx, endIdx, pendingComments);
      } else {
        // Single scalar value
        ScalarNode node = new ScalarNode(Tag.STR, firstValue, null, null, ScalarStyle.PLAIN);
        if (!pendingComments.isEmpty()) {
          node.setBlockComments(pendingComments);
        }
        return node;
      }
    }
  }

  private MappingNode parseMapping(List<Row> rows, int indentLevel, int startIdx, int endIdx,
      List<CommentLine> leadingComments) {
    List<NodeTuple> tuples = new ArrayList<>();
    int cellOffset = indentLevel * workbookSymbol.getIndentationCellNum();
    List<CommentLine> pendingComments = new ArrayList<>(leadingComments);

    int i = startIdx;
    while (i < endIdx) {
      Row row = rows.get(i);
      int rowIndent = getIndentLevel(row);

      if (rowIndent < indentLevel) {
        break; // Back to parent level
      }

      if (rowIndent > indentLevel) {
        i++;
        continue; // Skip nested content (handled by recursive calls)
      }

      String keyValue = getCellValue(row, cellOffset);
      if (keyValue == null) {
        i++;
        continue;
      }

      if (isComment(keyValue)) {
        pendingComments.add(createCommentLine(keyValue));
        i++;
        continue;
      }

      if (isItemMark(keyValue)) {
        break; // This is a sequence, not a mapping
      }

      // Create key node
      ScalarNode keyNode = new ScalarNode(Tag.STR, keyValue, null, null, ScalarStyle.PLAIN);
      if (!pendingComments.isEmpty()) {
        keyNode.setBlockComments(new ArrayList<>(pendingComments));
        pendingComments.clear();
      }

      // Check for key inline comment and inline value
      // Format can be: key | value | value_comment OR key | key_comment | value | value_comment
      String secondCell = getCellValue(row, cellOffset + 1);
      Node valueNode;

      if (secondCell != null) {
        int valueOffset;
        if (isComment(secondCell)) {
          // Second cell is a key inline comment: key | key_comment | value | value_comment
          List<CommentLine> keyInlineComments = new ArrayList<>();
          keyInlineComments.add(createInlineCommentLine(secondCell));
          keyNode.setInLineComments(keyInlineComments);
          valueOffset = cellOffset + 2;
        } else {
          // Second cell is the value: key | value | value_comment
          valueOffset = cellOffset + 1;
        }

        String inlineValue = getCellValue(row, valueOffset);
        if (inlineValue != null) {
          // Inline scalar value
          valueNode = new ScalarNode(Tag.STR, inlineValue, null, null, ScalarStyle.PLAIN);
          // Check for value inline comments
          List<CommentLine> inlineComments = parseInlineComments(row, valueOffset + 1);
          if (!inlineComments.isEmpty()) {
            valueNode.setInLineComments(inlineComments);
          }
          i++;
        } else {
          // Key with key inline comment but nested content
          int nestedStart = i + 1;
          int nestedEnd = findNestedEnd(rows, indentLevel, nestedStart, endIdx);
          valueNode = parseRows(rows, indentLevel + 1, nestedStart, nestedEnd);
          if (valueNode == null) {
            valueNode = new ScalarNode(Tag.STR, "", null, null, ScalarStyle.PLAIN);
          }
          i = nestedEnd;
        }
      } else {
        // Nested content - find extent
        int nestedStart = i + 1;
        int nestedEnd = findNestedEnd(rows, indentLevel, nestedStart, endIdx);
        valueNode = parseRows(rows, indentLevel + 1, nestedStart, nestedEnd);
        if (valueNode == null) {
          valueNode = new ScalarNode(Tag.STR, "", null, null, ScalarStyle.PLAIN);
        }
        i = nestedEnd;
      }

      tuples.add(new NodeTuple(keyNode, valueNode));
    }

    MappingNode node = new MappingNode(Tag.MAP, tuples, FlowStyle.BLOCK);
    return node;
  }

  private SequenceNode parseSequence(List<Row> rows, int indentLevel, int startIdx, int endIdx,
      List<CommentLine> leadingComments) {
    List<Node> items = new ArrayList<>();
    int cellOffset = indentLevel * workbookSymbol.getIndentationCellNum();
    List<CommentLine> pendingComments = new ArrayList<>(leadingComments);

    int i = startIdx;
    while (i < endIdx) {
      Row row = rows.get(i);
      int rowIndent = getIndentLevel(row);

      if (rowIndent < indentLevel) {
        break;
      }

      if (rowIndent > indentLevel) {
        i++;
        continue;
      }

      String firstValue = getCellValue(row, cellOffset);
      if (firstValue == null) {
        i++;
        continue;
      }

      if (isComment(firstValue)) {
        pendingComments.add(createCommentLine(firstValue));
        i++;
        continue;
      }

      if (!isItemMark(firstValue)) {
        break; // Not a sequence item
      }

      // Parse sequence item
      String inlineValue = getCellValue(row, cellOffset + 1);
      Node itemNode;

      if (inlineValue != null) {
        // Inline scalar value
        itemNode = new ScalarNode(Tag.STR, inlineValue, null, null, ScalarStyle.PLAIN);
        // Check for inline comments
        List<CommentLine> inlineComments = parseInlineComments(row, cellOffset + 2);
        if (!inlineComments.isEmpty()) {
          itemNode.setInLineComments(inlineComments);
        }
        i++;
      } else {
        // Nested content
        int nestedStart = i + 1;
        int nestedEnd = findNestedEnd(rows, indentLevel, nestedStart, endIdx);
        itemNode = parseRows(rows, indentLevel + 1, nestedStart, nestedEnd);
        if (itemNode == null) {
          itemNode = new ScalarNode(Tag.STR, "", null, null, ScalarStyle.PLAIN);
        }
        i = nestedEnd;
      }

      if (!pendingComments.isEmpty()) {
        itemNode.setBlockComments(new ArrayList<>(pendingComments));
        pendingComments.clear();
      }

      items.add(itemNode);
    }

    return new SequenceNode(Tag.SEQ, items, FlowStyle.BLOCK);
  }

  private int findNestedEnd(List<Row> rows, int parentIndent, int startIdx, int endIdx) {
    for (int i = startIdx; i < endIdx; i++) {
      Row row = rows.get(i);
      int rowIndent = getIndentLevel(row);
      if (rowIndent <= parentIndent) {
        return i;
      }
    }
    return endIdx;
  }

  private boolean hasNestedContent(List<Row> rows, int indentLevel, int startIdx, int endIdx) {
    if (startIdx + 1 >= endIdx) return false;
    Row nextRow = rows.get(startIdx + 1);
    int nextIndent = getIndentLevel(nextRow);
    return nextIndent > indentLevel;
  }

  private int getIndentLevel(Row row) {
    if (row == null) return 0;
    for (int i = 0; i <= row.getLastCellNum(); i++) {
      String value = getCellValue(row, i);
      if (value != null && !value.isEmpty()) {
        return i / workbookSymbol.getIndentationCellNum();
      }
    }
    return 0;
  }

  private String getCellValue(Row row, int cellIndex) {
    if (row == null || cellIndex < 0) return null;
    Cell cell = row.getCell(cellIndex);
    if (cell == null) return null;

    if (cell.getCellType() == CellType.STRING) {
      String value = cell.getStringCellValue();
      return (value == null || value.isEmpty()) ? null : value;
    } else if (cell.getCellType() == CellType.NUMERIC) {
      double numValue = cell.getNumericCellValue();
      if (numValue == (long) numValue) {
        return String.valueOf((long) numValue);
      }
      return String.valueOf(numValue);
    } else if (cell.getCellType() == CellType.BOOLEAN) {
      return String.valueOf(cell.getBooleanCellValue());
    }
    return null;
  }

  private boolean isComment(String value) {
    return value != null && value.startsWith(workbookSymbol.getCommentMark());
  }

  private boolean isItemMark(String value) {
    return workbookSymbol.getItemMark().equals(value);
  }

  private CommentLine createCommentLine(String commentValue) {
    String text = commentValue;
    if (text.startsWith(workbookSymbol.getCommentMark())) {
      text = text.substring(workbookSymbol.getCommentMark().length()).trim();
    }
    return new CommentLine(null, null, " " + text, CommentType.BLOCK);
  }

  private CommentLine createInlineCommentLine(String commentValue) {
    String text = commentValue;
    if (text.startsWith(workbookSymbol.getCommentMark())) {
      text = text.substring(workbookSymbol.getCommentMark().length()).trim();
    }
    return new CommentLine(null, null, " " + text, CommentType.IN_LINE);
  }

  private List<CommentLine> parseInlineComments(Row row, int startCellIndex) {
    List<CommentLine> comments = new ArrayList<>();
    for (int i = startCellIndex; i <= row.getLastCellNum(); i++) {
      String value = getCellValue(row, i);
      if (value != null && isComment(value)) {
        String text = value.substring(workbookSymbol.getCommentMark().length()).trim();
        comments.add(new CommentLine(null, null, " " + text, CommentType.IN_LINE));
      }
    }
    return comments;
  }

}
