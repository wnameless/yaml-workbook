package com.github.wnameless.workbook.yamlworkbook;

import java.util.function.Function;

public interface SheetNameStrategy extends Function<Integer, String> {

  public static final SheetNameStrategy DEFAULT = new DefaultSheetNameStrategy();

  /**
   * Returns the hidden sheet name for a given visible sheet index. By default, appends "Hidden" to
   * the visible sheet name.
   *
   * @param visibleSheetIndex the index of the visible sheet
   * @return the hidden sheet name
   */
  default String applyHidden(int visibleSheetIndex) {
    return apply(visibleSheetIndex) + "Hidden";
  }

}
