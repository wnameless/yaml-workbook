package com.github.wnameless.workbook.yamlworkbook;

import java.util.function.Function;

public interface SheetNameStrategy extends Function<Integer, String> {

  public static final SheetNameStrategy DEFAULT = new DefaultSheetNameStrategy();

}
