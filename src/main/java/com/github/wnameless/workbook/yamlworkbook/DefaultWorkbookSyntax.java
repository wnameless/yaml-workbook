package com.github.wnameless.workbook.yamlworkbook;

import lombok.Data;

/**
 * Default implementation of {@link WorkbookSyntax} with standard YAML symbols.
 *
 * @author Wei-Ming Wu
 */
@Data
public final class DefaultWorkbookSyntax implements WorkbookSyntax {

  public final String frontmatter = "---";
  public final String commentMark = "#";
  public final String valueEscapeMark = "\\";
  public final String itemMark = "-";
  public final Short indentationCellNum = 1;

}
