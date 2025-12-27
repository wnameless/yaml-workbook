package com.github.wnameless.workbook.yamlworkbook;

import java.util.function.Function;

/**
 * Strategy for generating sheet names in workbooks.
 * <p>
 * Provides naming for both visible sheets and hidden sheets (used for large enum dropdowns).
 *
 * @author Wei-Ming Wu
 */
public interface SheetNameStrategy extends Function<Integer, String> {

  /** Default implementation using "Sheet1", "Sheet2", etc. naming pattern. */
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
