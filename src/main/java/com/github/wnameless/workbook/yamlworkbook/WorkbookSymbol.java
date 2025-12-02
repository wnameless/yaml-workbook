package com.github.wnameless.workbook.yamlworkbook;

public interface WorkbookSymbol {

  public static final WorkbookSymbol DEFAULT = new DefaultWorkbookSymbol();

  String getFrontmatter();

  String getCommentMark();

  String getValueEscapeMark();

  String getItemMark();

  Short getIndentationCellNum();

}
