package com.github.wnameless.workbook.yamlworkbook;

/**
 * Default implementation of {@link SheetNameStrategy} using "Sheet1", "Sheet2", etc. naming
 * pattern.
 *
 * @author Wei-Ming Wu
 */
public final class DefaultSheetNameStrategy implements SheetNameStrategy {

  @Override
  public String apply(Integer t) {
    return "Sheet" + (t == null ? 1 : t + 1);
  }

}
