package com.github.wnameless.workbook.yamlworkbook;

public interface WorkbookSyntax {

  public static final WorkbookSyntax DEFAULT = new DefaultWorkbookSyntax();

  String getFrontmatter();

  String getCommentMark();

  String getValueEscapeMark();

  String getItemMark();

  Short getIndentationCellNum();

}
