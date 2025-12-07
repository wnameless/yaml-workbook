package com.github.wnameless.workbook.yamlworkbook;

public final class DefaultSheetNameStrategy implements SheetNameStrategy {

  @Override
  public String apply(Integer t) {
    return "Sheet" + (t == null ? 1 : t + 1);
  }

}
