package com.github.wnameless.workbook.yamlworkbook;

import lombok.Data;

@Data
public final class DefaultWorkbookSymbol implements WorkbookSymbol {

  public final String frontmatter = "---";
  public final String commentMark = "#";
  public final String itemMark = "-";
  public final Short indentationCellNum = 1;

}
