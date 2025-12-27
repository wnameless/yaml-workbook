package com.github.wnameless.workbook.yamlworkbook;

import java.io.Reader;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.logging.Logger;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.ss.usermodel.DataValidationConstraint;
import org.apache.poi.ss.usermodel.DataValidationHelper;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Name;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.yaml.snakeyaml.LoaderOptions;
import org.yaml.snakeyaml.Yaml;
import org.yaml.snakeyaml.comments.CommentLine;
import org.yaml.snakeyaml.nodes.MappingNode;
import org.yaml.snakeyaml.nodes.Node;
import org.yaml.snakeyaml.nodes.NodeTuple;
import org.yaml.snakeyaml.nodes.ScalarNode;
import org.yaml.snakeyaml.nodes.SequenceNode;
import com.github.wnameless.json.jsonschemadatagenerator.AllOfOption;
import com.github.wnameless.json.jsonschemadatagenerator.JsonSchemaDataGenerator;
import com.github.wnameless.json.jsonschemadatagenerator.JsonSchemaPathNavigator;
import lombok.Builder;
import tools.jackson.databind.JsonNode;

@Builder
public class YamlWorkbookWriter {

  private static final Logger log = Logger.getLogger(YamlWorkbookWriter.class.getName());

  @Builder.Default
  private PrintMode printMode = PrintMode.YAML_ORIENTED;
  @Builder.Default
  private DisplayModeConfig displayModeConfig = DisplayModeConfig.DEFAULT;
  @Builder.Default
  private DataCollectConfig dataCollectConfig = DataCollectConfig.DEFAULT;
  @Builder.Default
  private WorkbookSyntax workbookSyntax = WorkbookSyntax.DEFAULT;
  @Builder.Default
  private NodeToSheetMapper nodeToSheetMapper = NodeToSheetMapper.DEFAULT;
  @Builder.Default
  private SheetNameStrategy sheetNameStrategy = SheetNameStrategy.DEFAULT;

  /** JSON Schema string for DATA_COLLECT mode */
  private String jsonSchema;

  // Internal state for sheet tracking (not part of builder)
  private final List<Sheet> visibleSheets = new ArrayList<>();
  private final Map<Integer, Sheet> hiddenSheets = new HashMap<>();
  private final Map<Integer, Integer> hiddenSheetEnumRowCounter = new HashMap<>();

  public Workbook toWorkbook(Reader yamlContent, Reader... yamlContents) {
    var workbook = new XSSFWorkbook();

    LoaderOptions options = new LoaderOptions();
    options.setProcessComments(true);
    Yaml yaml = new Yaml(options);

    List<Iterable<Node>> nodeIters = new ArrayList<>();
    nodeIters.add(yaml.composeAll(yamlContent));
    for (Reader content : yamlContents) {
      nodeIters.add(yaml.composeAll(content));
    }

    processNodes(nodeIters, workbook);

    if (visibleSheets.isEmpty()) {
      Sheet sheet = createVisibleSheet(workbook, 0);
      visibleSheets.add(sheet);
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
    var logicalSheetIdx = nodeToSheetMapper.apply(node, nodeIdx);

    // Ensure visible sheet exists at logical index
    while (visibleSheets.size() <= logicalSheetIdx) {
      Sheet newSheet = createVisibleSheet(workbook, visibleSheets.size());
      visibleSheets.add(newSheet);
    }
    var sheet = visibleSheets.get(logicalSheetIdx);

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

  private Sheet createVisibleSheet(Workbook workbook, int logicalIdx) {
    String sheetName = sheetNameStrategy.apply(logicalIdx);
    return workbook.createSheet(sheetName);
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
        String commentText = extractCommentText(comments);
        cell.setCellValue(commentText);
        // Store original comment with # prefix for roundtrip support
        addCellComment(cell, workbookSyntax.getCommentMark() + " " + commentText);
      }
      case HIDDEN -> {
        /* skip comments entirely */ }
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
        String originalKeyValue = null;
        if (isDisplayMode()) {
          String keyCommentText = extractCommentText(scalarKey.getInLineComments());
          if (keyCommentText != null) {
            switch (displayModeConfig.getKeyComment()) {
              case DISPLAY_NAME -> {
                originalKeyValue = scalarKey.getValue();
                keyDisplayValue = keyCommentText;
              }
              case HIDDEN -> {
                /* keep original */ }
              case COMMENT -> {
                /* handled separately */ }
            }
          }
        }
        keyCell.setCellValue(keyDisplayValue);
        if (originalKeyValue != null) {
          addCellComment(keyCell, originalKeyValue);
        }

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
          String originalValue = null;
          if (isDisplayMode()) {
            String valueCommentText = extractCommentText(scalarValue.getInLineComments());
            if (valueCommentText != null) {
              switch (displayModeConfig.getValueComment()) {
                case DISPLAY_NAME -> {
                  originalValue = scalarValue.getValue();
                  valueDisplayValue = valueCommentText;
                }
                case HIDDEN -> {
                  /* keep original */ }
                case COMMENT -> {
                  /* handled separately */ }
              }
            }
          }
          valueCell.setCellValue(valueDisplayValue);
          if (originalValue != null) {
            addCellComment(valueCell, originalValue);
          }

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
    if (value == null || "null".equals(value)) {
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
    return printMode == PrintMode.WORKBOOK_READABLE;
  }

  private String extractCommentText(List<CommentLine> comments) {
    if (comments == null || comments.isEmpty()) {
      return null;
    }
    // Use the first comment's text as display name
    return comments.get(0).getValue().trim();
  }

  private void addCellComment(Cell cell, String commentText) {
    if (commentText == null || commentText.isEmpty()) {
      return;
    }
    Sheet sheet = cell.getSheet();
    Workbook workbook = sheet.getWorkbook();
    CreationHelper factory = workbook.getCreationHelper();
    Drawing<?> drawing = sheet.createDrawingPatriarch();

    ClientAnchor anchor = factory.createClientAnchor();
    anchor.setCol1(cell.getColumnIndex());
    anchor.setCol2(cell.getColumnIndex() + 2);
    anchor.setRow1(cell.getRowIndex());
    anchor.setRow2(cell.getRowIndex() + 2);

    Comment comment = drawing.createCellComment(anchor);
    comment.setString(factory.createRichTextString(commentText));
    cell.setCellComment(comment);
  }

  // ==================== DATA_COLLECT Mode Methods ====================

  /**
   * Creates a workbook from JSON Schema for DATA_COLLECT mode.
   *
   * @return the generated workbook with dropdowns and schema metadata
   * @throws IllegalStateException if not in DATA_COLLECT mode or jsonSchema is null
   * @throws RuntimeException if schema parsing or processing fails
   */
  public Workbook toWorkbook() {
    if (printMode != PrintMode.DATA_COLLECT || jsonSchema == null) {
      throw new IllegalStateException(
          "toWorkbook() without parameters requires DATA_COLLECT mode and jsonSchema to be set");
    }

    try {
      var workbook = new XSSFWorkbook();

      // 1. Generate skeleton JSON from schema
      var generator = JsonSchemaDataGenerator.skeleton();
      if (dataCollectConfig.isSkipAllOf()) {
        generator = generator.withAllOfOption(AllOfOption.SKIP);
      }
      JsonNode skeleton = generator.generate(jsonSchema);

      // 2. Create navigator for metadata lookup
      JsonSchemaPathNavigator navigator = JsonSchemaPathNavigator.of(jsonSchema);

      // 3. Convert to YAML Node
      Node yamlNode = JsonNodeToYamlNodeConverter.convert(skeleton);

      // 4. Create visible sheet and process with path tracking
      Sheet sheet = createVisibleSheet(workbook, 0);
      visibleSheets.add(sheet);

      // Write frontmatter
      writeFrontmatter(sheet);

      // Process the node with path tracking
      traverseAndPrintNodeWithPath(yamlNode, sheet, 0, "$", navigator);

      return workbook;
    } catch (Exception e) {
      throw new RuntimeException("Failed to generate workbook from JSON Schema", e);
    }
  }

  private void traverseAndPrintNodeWithPath(Node node, Sheet sheet, int indentLevel,
      String jsonPath, JsonSchemaPathNavigator navigator) {
    if (node == null) {
      return;
    }

    if (node instanceof ScalarNode scalarNode) {
      traverseScalarNodeWithPath(scalarNode, sheet, indentLevel, jsonPath, navigator);
    } else if (node instanceof MappingNode mappingNode) {
      traverseMappingNodeWithPath(mappingNode, sheet, indentLevel, jsonPath, navigator);
    } else if (node instanceof SequenceNode sequenceNode) {
      traverseSequenceNodeWithPath(sequenceNode, sheet, indentLevel, jsonPath, navigator);
    }
  }

  private void traverseScalarNodeWithPath(ScalarNode node, Sheet sheet, int indentLevel,
      String jsonPath, JsonSchemaPathNavigator navigator) {
    Row row = sheet.createRow(sheet.getLastRowNum() + 1);
    int cellIndex = indentLevel * workbookSyntax.getIndentationCellNum();
    Cell cell = row.createCell(cellIndex);

    String value = node.getValue();
    JsonNode schema = navigator.findSchema(jsonPath).orElse(null);

    // Handle enum with enumNames
    if (schema != null && schema.has("enum")) {
      handleEnumCell(cell, schema, sheet);
    } else {
      cell.setCellValue(escapeValueIfNeeded(value));
    }
  }

  private void traverseMappingNodeWithPath(MappingNode node, Sheet sheet, int indentLevel,
      String jsonPath, JsonSchemaPathNavigator navigator) {
    for (NodeTuple tuple : node.getValue()) {
      Node keyNode = tuple.getKeyNode();
      Node valueNode = tuple.getValueNode();

      if (keyNode instanceof ScalarNode scalarKey) {
        String originalKey = scalarKey.getValue();
        String propertyPath =
            "$".equals(jsonPath) ? "$." + originalKey : jsonPath + "." + originalKey;
        JsonNode propertySchema = navigator.findSchema(propertyPath).orElse(null);

        Row row = sheet.createRow(sheet.getLastRowNum() + 1);
        int cellIndex = indentLevel * workbookSyntax.getIndentationCellNum();
        Cell keyCell = row.createCell(cellIndex);

        // Use title as display name if available
        String displayKey = getDisplayNameForKey(originalKey, propertySchema);
        keyCell.setCellValue(displayKey);

        // Store original key in comment if title was used
        if (shouldStoreOriginalKey(propertySchema)) {
          addCellComment(keyCell, originalKey);
        }

        if (valueNode instanceof ScalarNode scalarValue) {
          int nextCellIndex = cellIndex + 1;
          Cell valueCell = row.createCell(nextCellIndex);

          // Handle enum with enumNames for value
          if (propertySchema != null && propertySchema.has("enum")) {
            handleEnumCell(valueCell, propertySchema, sheet);
          } else {
            valueCell.setCellValue(escapeValueIfNeeded(scalarValue.getValue()));
          }
        } else {
          traverseAndPrintNodeWithPath(valueNode, sheet, indentLevel + 1, propertyPath, navigator);
        }
      } else {
        traverseAndPrintNodeWithPath(keyNode, sheet, indentLevel, jsonPath, navigator);
        traverseAndPrintNodeWithPath(valueNode, sheet, indentLevel + 1, jsonPath, navigator);
      }
    }
  }

  private void traverseSequenceNodeWithPath(SequenceNode node, Sheet sheet, int indentLevel,
      String jsonPath, JsonSchemaPathNavigator navigator) {
    String itemsPath = jsonPath + "[*]";

    for (Node item : node.getValue()) {
      Row row = sheet.createRow(sheet.getLastRowNum() + 1);
      int cellIndex = indentLevel * workbookSyntax.getIndentationCellNum();
      Cell itemMarkCell = row.createCell(cellIndex);
      itemMarkCell.setCellValue(workbookSyntax.getItemMark());

      if (item instanceof ScalarNode scalarItem) {
        Cell valueCell = row.createCell(cellIndex + 1);
        JsonNode itemSchema = navigator.findSchema(itemsPath).orElse(null);

        // Handle enum with enumNames for array items
        if (itemSchema != null && itemSchema.has("enum")) {
          handleEnumCell(valueCell, itemSchema, sheet);
        } else {
          valueCell.setCellValue(escapeValueIfNeeded(scalarItem.getValue()));
        }
      } else {
        traverseAndPrintNodeWithPath(item, sheet, indentLevel + 1, itemsPath, navigator);
      }
    }
  }

  // ==================== Helper Methods for DATA_COLLECT ====================

  private void handleEnumCell(Cell cell, JsonNode schema, Sheet sheet) {
    JsonNode enumValues = schema.get("enum");
    JsonNode enumNames = schema.has("enumNames") ? schema.get("enumNames") : null;

    if (enumNames != null) {
      // Display enumNames in dropdown
      List<String> displayOptions = toStringList(enumNames);
      addDropdownValidation(cell, displayOptions, sheet);

      // Store ENUM_VALUES in cell comment for roundtrip (index-based lookup)
      String valuesComment = buildEnumValues(enumValues);
      addCellComment(cell, valuesComment);
    } else {
      // Use enum values directly as dropdown
      List<String> options = toJsonStringList(enumValues);
      addDropdownValidation(cell, options, sheet);
      // No comment needed - dropdown values ARE the actual values
    }
  }

  private String getDisplayNameForKey(String originalKey, JsonNode schema) {
    if (schema != null && schema.has("title")) {
      return schema.get("title").asString();
    }
    return originalKey;
  }

  private boolean shouldStoreOriginalKey(JsonNode schema) {
    return schema != null && schema.has("title");
  }

  private void addDropdownValidation(Cell cell, List<String> options, Sheet sheet) {
    if (options == null || options.isEmpty()) {
      return;
    }

    String joinedOptions = String.join(",", options);

    if (joinedOptions.length() <= 255) {
      // Use explicit list constraint (current behavior)
      addExplicitDropdownValidation(cell, options, sheet);
    } else if (dataCollectConfig.isUseHiddenSheetsForLongEnums()) {
      // Write to hidden sheet, use named range
      addNamedRangeDropdownValidation(cell, options, sheet);
    } else {
      // Truncate + warning
      List<String> truncated = truncateOptionsTo256(options);
      log.warning(String.format("Dropdown truncated from %d to %d options (256 char limit)",
          options.size(), truncated.size()));
      addExplicitDropdownValidation(cell, truncated, sheet);
    }
  }

  private void addExplicitDropdownValidation(Cell cell, List<String> options, Sheet sheet) {
    DataValidationHelper validationHelper = sheet.getDataValidationHelper();
    DataValidationConstraint constraint =
        validationHelper.createExplicitListConstraint(options.toArray(new String[0]));
    CellRangeAddressList addressList = new CellRangeAddressList(cell.getRowIndex(),
        cell.getRowIndex(), cell.getColumnIndex(), cell.getColumnIndex());
    DataValidation validation = validationHelper.createValidation(constraint, addressList);
    // For XSSF, setSuppressDropDownArrow(true) actually SHOWS the dropdown arrow
    validation.setSuppressDropDownArrow(true);
    sheet.addValidationData(validation);
  }

  private void addNamedRangeDropdownValidation(Cell cell, List<String> options, Sheet sheet) {
    Workbook workbook = sheet.getWorkbook();
    int visibleSheetIdx = workbook.getSheetIndex(sheet);

    // Get or create hidden sheet (lazy)
    Sheet hiddenSheet = getOrCreateHiddenSheet(workbook, visibleSheetIdx);

    // Write options to hidden sheet column
    int startRow = hiddenSheetEnumRowCounter.getOrDefault(visibleSheetIdx, 0);
    for (int i = 0; i < options.size(); i++) {
      Row row = hiddenSheet.getRow(startRow + i);
      if (row == null) {
        row = hiddenSheet.createRow(startRow + i);
      }
      row.createCell(0).setCellValue(options.get(i));
    }
    hiddenSheetEnumRowCounter.put(visibleSheetIdx, startRow + options.size());

    // Create named range
    String rangeName = "Enum_" + cell.getRowIndex() + "_" + cell.getColumnIndex();
    Name namedRange = workbook.createName();
    namedRange.setNameName(rangeName);
    String formula = String.format("'%s'!$A$%d:$A$%d", hiddenSheet.getSheetName(), startRow + 1,
        startRow + options.size());
    namedRange.setRefersToFormula(formula);

    // Create validation using named range
    DataValidationHelper helper = sheet.getDataValidationHelper();
    DataValidationConstraint constraint = helper.createFormulaListConstraint(rangeName);
    CellRangeAddressList addressList = new CellRangeAddressList(cell.getRowIndex(),
        cell.getRowIndex(), cell.getColumnIndex(), cell.getColumnIndex());
    DataValidation validation = helper.createValidation(constraint, addressList);
    validation.setSuppressDropDownArrow(true);
    sheet.addValidationData(validation);
  }

  private Sheet getOrCreateHiddenSheet(Workbook workbook, int visibleLogicalIdx) {
    if (hiddenSheets.containsKey(visibleLogicalIdx)) {
      return hiddenSheets.get(visibleLogicalIdx);
    }

    String hiddenSheetName = sheetNameStrategy.applyHidden(visibleLogicalIdx);

    // Create hidden sheet
    Sheet hiddenSheet = workbook.createSheet(hiddenSheetName);

    // Find actual index of visible sheet and insert hidden sheet right after it
    Sheet visibleSheet = visibleSheets.get(visibleLogicalIdx);
    int visibleActualIdx = workbook.getSheetIndex(visibleSheet);
    workbook.setSheetOrder(hiddenSheetName, visibleActualIdx + 1);
    workbook.setSheetHidden(visibleActualIdx + 1, true);

    hiddenSheets.put(visibleLogicalIdx, hiddenSheet);
    hiddenSheetEnumRowCounter.put(visibleLogicalIdx, 0);
    return hiddenSheet;
  }

  private List<String> truncateOptionsTo256(List<String> options) {
    List<String> result = new ArrayList<>();
    int totalLength = 0;
    for (String opt : options) {
      int addedLength = (result.isEmpty() ? 0 : 1) + opt.length(); // +1 for comma
      if (totalLength + addedLength > 255) {
        break;
      }
      result.add(opt);
      totalLength += addedLength;
    }
    return result;
  }

  private String buildEnumValues(JsonNode enumValues) {
    StringBuilder sb = new StringBuilder("ENUM_VALUES:");
    for (int i = 0; i < enumValues.size(); i++) {
      if (i > 0) {
        sb.append(",");
      }
      sb.append(escapeEnumPart(toJsonValue(enumValues.get(i))));
    }
    return sb.toString();
  }

  private String escapeEnumPart(String s) {
    if (s == null) {
      return "";
    }
    // Only need to escape comma and backslash (no longer need = since we use index-based lookup)
    return s.replace("\\", "\\\\").replace(",", "\\,");
  }

  private String toJsonValue(JsonNode node) {
    if (node == null || node.isNull()) {
      return "null";
    }
    if (node.isNumber() || node.isBoolean()) {
      return node.asString();
    }
    // For strings and other types
    return node.asString();
  }

  private List<String> toStringList(JsonNode arrayNode) {
    List<String> result = new ArrayList<>();
    if (arrayNode != null && arrayNode.isArray()) {
      for (JsonNode item : arrayNode) {
        result.add(item.asString());
      }
    }
    return result;
  }

  private List<String> toJsonStringList(JsonNode arrayNode) {
    List<String> result = new ArrayList<>();
    if (arrayNode != null && arrayNode.isArray()) {
      for (JsonNode item : arrayNode) {
        result.add(toJsonValue(item));
      }
    }
    return result;
  }

}
