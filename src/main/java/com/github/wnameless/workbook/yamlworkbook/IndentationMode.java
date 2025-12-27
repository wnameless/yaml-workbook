package com.github.wnameless.workbook.yamlworkbook;

/**
 * Defines how indentation is represented in the workbook.
 */
public enum IndentationMode {

  /**
   * Uses empty cells for indentation (default behavior). Cell index = indentLevel *
   * indentationCellNum. Suitable for most use cases with shallow nesting.
   */
  CELL_OFFSET,

  /**
   * Uses a prefix marker in a separate cell to indicate indent level. Layout: [prefix] [key]
   * [value]. Level 0 has no prefix (content at col 0), levels 1+ have prefix at col 0 and content
   * at col 1. Suitable for deeply nested structures to avoid horizontal scrolling.
   */
  PREFIX

}
