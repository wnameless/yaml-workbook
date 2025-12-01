package com.github.wnameless.workbook.yamlworkbook;

public final class DefaultWorkbookSheetNameStrategy implements WorkbookSheetNameStrategy {

  @Override
  public String apply(Integer t) {
    return "Sheet" + (t == null ? 1 : t + 1);
  }

}
