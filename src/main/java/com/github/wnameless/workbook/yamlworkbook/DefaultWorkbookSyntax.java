package com.github.wnameless.workbook.yamlworkbook;

import lombok.Data;

@Data
public final class DefaultWorkbookSyntax implements WorkbookSyntax {

  public final String frontmatter = "---";
  public final String commentMark = "#";
  public final String valueEscapeMark = "\\";
  public final String itemMark = "-";
  public final Short indentationCellNum = 1;

}
