package com.github.wnameless.workbook.yamlworkbook;

/**
 * Defines the syntax symbols used for YAML/workbook conversion.
 * <p>
 * Provides configurable markers for frontmatter, comments, value escaping, sequence items, and
 * indentation cell count.
 *
 * @author Wei-Ming Wu
 */
public interface WorkbookSyntax {

  /** Default implementation with standard YAML symbols. */
  public static final WorkbookSyntax DEFAULT = new DefaultWorkbookSyntax();

  /**
   * Returns the frontmatter marker (document separator).
   *
   * @return the frontmatter marker, typically "---"
   */
  String getFrontmatter();

  /**
   * Returns the comment marker prefix.
   *
   * @return the comment marker, typically "#"
   */
  String getCommentMark();

  /**
   * Returns the escape marker for values starting with special characters.
   *
   * @return the escape marker, typically "\\"
   */
  String getValueEscapeMark();

  /**
   * Returns the sequence item marker.
   *
   * @return the item marker, typically "-"
   */
  String getItemMark();

  /**
   * Returns the number of cells used per indentation level.
   *
   * @return the indentation cell count, typically 1
   */
  Short getIndentationCellNum();

}

