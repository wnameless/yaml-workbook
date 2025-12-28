package com.github.wnameless.workbook.yamlworkbook;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.List;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.ss.usermodel.DataValidationConstraint;
import org.apache.poi.ss.usermodel.Name;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;
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

/**
 * Converts Excel workbooks back to SnakeYAML Node trees for roundtrip support.
 * <p>
 * Uses Lombok's {@code @Builder} pattern for configuration. Key features:
 * <ul>
 * <li>Reconstructs SnakeYAML Node trees from workbook cells</li>
 * <li>Preserves structure and comments for roundtrip conversion</li>
 * <li>Supports the same output modes and indentation modes as {@link YamlWorkbookWriter}</li>
 * </ul>
 *
 * @author Wei-Ming Wu
 * @see YamlWorkbookWriter
 * @see OutputMode
 * @see IndentationMode
 */
@Builder
public class YamlWorkbookReader {

  @Builder.Default
  private OutputMode outputMode = OutputMode.YAML_ORIENTED;
  @Builder.Default
  private WorkbookSyntax workbookSyntax = WorkbookSyntax.DEFAULT;
  @Builder.Default
  private SheetNameStrategy sheetNameStrategy = SheetNameStrategy.DEFAULT;
  @Builder.Default
  private IndentationMode indentationMode = IndentationMode.CELL_OFFSET;
  @Builder.Default
  private IndentPrefixStrategy indentPrefixStrategy = IndentPrefixStrategy.DEFAULT;

  /**
   * Converts an Excel workbook to a list of SnakeYAML Node objects.
   *
   * @param workbook the Excel workbook to convert (may be null)
   * @return a list of YAML document nodes, or empty list if workbook is null
   */
  public List<Node> fromWorkbook(Workbook workbook) {
    var nodeList = new ArrayList<Node>();
    if (workbook == null) return nodeList;

    // Build list of visible sheets (skip hidden sheets)
    List<Sheet> visibleSheets = new ArrayList<>();
    for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
      if (!workbook.isSheetHidden(i)) {
        visibleSheets.add(workbook.getSheetAt(i));
      }
    }

    // Process visible sheets by logical index
    for (int logicalIdx = 0; logicalIdx < visibleSheets.size(); logicalIdx++) {
      Sheet sheet = visibleSheets.get(logicalIdx);
      String expectedName = sheetNameStrategy.apply(logicalIdx);
      if (sheet.getSheetName().equals(expectedName)) {
        processYamlSheet(sheet).forEach(nodeList::add);
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
      if (workbookSyntax.getFrontmatter().equals(firstCellValue)) {
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
    int cellOffset = getContentOffset(indentLevel);

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
        ScalarNode node = new ScalarNode(Tag.STR, unescapeValueIfNeeded(firstValue), null, null,
            ScalarStyle.PLAIN);
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
    int cellOffset = getContentOffset(indentLevel);
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
          valueNode = new ScalarNode(Tag.STR, unescapeValueIfNeeded(inlineValue), null, null,
              ScalarStyle.PLAIN);
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
    int cellOffset = getContentOffset(indentLevel);
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
        itemNode = new ScalarNode(Tag.STR, unescapeValueIfNeeded(inlineValue), null, null,
            ScalarStyle.PLAIN);
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

  private boolean isPrefixMode() {
    return indentationMode == IndentationMode.PREFIX;
  }

  private int getContentOffset(int indentLevel) {
    if (isPrefixMode()) {
      // In prefix mode: level 0 content at col 0, levels 1+ content at col 1
      return indentLevel > 0 ? 1 : 0;
    }
    return indentLevel * workbookSyntax.getIndentCellCount();
  }

  private int getIndentLevel(Row row) {
    if (row == null) return 0;

    if (isPrefixMode()) {
      // In prefix mode, check cell 0 for a prefix
      String firstCell = getCellStringValue(row.getCell(0));
      if (firstCell == null || firstCell.isEmpty()) {
        // No prefix means level 0 (or empty row)
        // Check if there's content at col 0
        String content = getCellValue(row, 0);
        return (content != null && !content.isEmpty()) ? 0 : 0;
      }
      int level = indentPrefixStrategy.parsePrefix(firstCell);
      if (level > 0) {
        return level;
      }
      // Not a valid prefix, treat as level 0 content
      return 0;
    }

    // Original CELL_OFFSET behavior
    for (int i = 0; i <= row.getLastCellNum(); i++) {
      String value = getCellValue(row, i);
      if (value != null && !value.isEmpty()) {
        return i / workbookSyntax.getIndentCellCount();
      }
    }
    return 0;
  }

  private String getCellValue(Row row, int cellIndex) {
    if (row == null || cellIndex < 0) return null;
    Cell cell = row.getCell(cellIndex);
    if (cell == null) return null;

    // In DISPLAY_MODE or FORM_MODE, check cell comments for original values
    if (isReadableMode()) {
      String commentValue = getCellCommentValue(cell);
      if (commentValue != null) {
        if (commentValue.startsWith("ENUM_VALUES:")) {
          // Enum with enumNames: map display value back to actual enum value by index
          String displayValue = getCellStringValue(cell);
          List<String> dropdownOptions = getDropdownOptionsForCell(cell);
          return mapEnumValueByIndex(displayValue, dropdownOptions, commentValue);
        } else {
          // Cell comment contains the original value (or original comment with # prefix)
          return commentValue;
        }
      }
    }

    return getCellStringValue(cell);
  }

  private String getCellStringValue(Cell cell) {
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

  private String getCellCommentValue(Cell cell) {
    if (cell == null) return null;
    Comment comment = cell.getCellComment();
    if (comment == null) return null;
    String commentText = comment.getString().getString();
    return (commentText == null || commentText.isEmpty()) ? null : commentText;
  }

  private boolean isReadableMode() {
    return outputMode == OutputMode.DISPLAY_MODE || outputMode == OutputMode.FORM_MODE;
  }

  private List<String> getDropdownOptionsForCell(Cell cell) {
    if (cell == null) {
      return Collections.emptyList();
    }
    Sheet sheet = cell.getSheet();
    int row = cell.getRowIndex();
    int col = cell.getColumnIndex();

    for (DataValidation validation : sheet.getDataValidations()) {
      if (cellInRange(row, col, validation.getRegions().getCellRangeAddresses())) {
        DataValidationConstraint constraint = validation.getValidationConstraint();
        if (constraint.getValidationType() == DataValidationConstraint.ValidationType.LIST) {
          // Try explicit list first
          String[] explicitOptions = constraint.getExplicitListValues();
          if (explicitOptions != null) {
            return Arrays.asList(explicitOptions);
          }
          // Try formula-based constraint (named range)
          String formula = constraint.getFormula1();
          if (formula != null) {
            return getOptionsFromNamedRange(cell.getSheet().getWorkbook(), formula);
          }
        }
      }
    }
    return Collections.emptyList();
  }

  private List<String> getOptionsFromNamedRange(Workbook workbook, String rangeName) {
    Name name = workbook.getName(rangeName);
    if (name == null) {
      return Collections.emptyList();
    }

    String formula = name.getRefersToFormula();
    // Parse formula like 'Sheet1Hidden'!$A$1:$A$10
    try {
      AreaReference areaRef = new AreaReference(formula, workbook.getSpreadsheetVersion());
      Sheet sheet = workbook.getSheet(areaRef.getFirstCell().getSheetName());
      if (sheet == null) {
        return Collections.emptyList();
      }

      List<String> options = new ArrayList<>();
      CellReference first = areaRef.getFirstCell();
      CellReference last = areaRef.getLastCell();
      for (int r = first.getRow(); r <= last.getRow(); r++) {
        Row row = sheet.getRow(r);
        if (row != null) {
          Cell c = row.getCell(first.getCol());
          if (c != null) {
            options.add(c.getStringCellValue());
          }
        }
      }
      return options;
    } catch (Exception e) {
      return Collections.emptyList();
    }
  }

  private boolean cellInRange(int row, int col, CellRangeAddress[] ranges) {
    for (CellRangeAddress range : ranges) {
      if (row >= range.getFirstRow() && row <= range.getLastRow() && col >= range.getFirstColumn()
          && col <= range.getLastColumn()) {
        return true;
      }
    }
    return false;
  }

  private String mapEnumValueByIndex(String displayValue, List<String> dropdownOptions,
      String comment) {
    if (displayValue == null) {
      return null;
    }
    int index = dropdownOptions.indexOf(displayValue);
    if (index < 0) {
      return displayValue; // Fallback: return as-is if not found in dropdown
    }

    List<String> enumValues = parseEnumValues(comment);
    return index < enumValues.size() ? enumValues.get(index) : displayValue;
  }

  private List<String> parseEnumValues(String comment) {
    // Parse "ENUM_VALUES:val1,val2,val3" with escape handling for \, and \\
    String valuesPart = comment.substring("ENUM_VALUES:".length());
    List<String> result = new ArrayList<>();
    StringBuilder current = new StringBuilder();
    boolean escaped = false;

    for (char c : valuesPart.toCharArray()) {
      if (escaped) {
        current.append(c);
        escaped = false;
      } else if (c == '\\') {
        escaped = true;
      } else if (c == ',') {
        result.add(current.toString());
        current.setLength(0);
      } else {
        current.append(c);
      }
    }
    result.add(current.toString()); // Add last value
    return result;
  }

  private boolean isComment(String value) {
    return value != null && value.startsWith(workbookSyntax.getCommentMark())
        && !value.startsWith(workbookSyntax.getEscapeMark() + workbookSyntax.getCommentMark());
  }

  private String unescapeValueIfNeeded(String value) {
    if (value == null) {
      return null;
    }
    // Only unescape if value STARTS with escape mark
    if (value.startsWith(workbookSyntax.getEscapeMark())) {
      return value.substring(workbookSyntax.getEscapeMark().length());
    }
    return value;
  }

  private boolean isItemMark(String value) {
    return workbookSyntax.getItemMark().equals(value);
  }

  private CommentLine createCommentLine(String commentValue) {
    String text = commentValue;
    if (text.startsWith(workbookSyntax.getCommentMark())) {
      text = text.substring(workbookSyntax.getCommentMark().length()).trim();
    }
    return new CommentLine(null, null, " " + text, CommentType.BLOCK);
  }

  private CommentLine createInlineCommentLine(String commentValue) {
    String text = commentValue;
    if (text.startsWith(workbookSyntax.getCommentMark())) {
      text = text.substring(workbookSyntax.getCommentMark().length()).trim();
    }
    return new CommentLine(null, null, " " + text, CommentType.IN_LINE);
  }

  private List<CommentLine> parseInlineComments(Row row, int startCellIndex) {
    List<CommentLine> comments = new ArrayList<>();
    for (int i = startCellIndex; i <= row.getLastCellNum(); i++) {
      String value = getCellValue(row, i);
      if (value != null && isComment(value)) {
        String text = value.substring(workbookSyntax.getCommentMark().length()).trim();
        comments.add(new CommentLine(null, null, " " + text, CommentType.IN_LINE));
      }
    }
    return comments;
  }

  /**
   * Builder class for {@link YamlWorkbookReader}.
   * <p>
   * This stub class is completed by Lombok's {@code @Builder} annotation processor.
   */
  public static class YamlWorkbookReaderBuilder {}

}
