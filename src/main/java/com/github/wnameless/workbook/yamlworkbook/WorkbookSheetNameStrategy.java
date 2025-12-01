package com.github.wnameless.workbook.yamlworkbook;

import java.util.function.Function;

public interface WorkbookSheetNameStrategy extends Function<Integer, String> {

  public static final WorkbookSheetNameStrategy DEFAULT = new DefaultWorkbookSheetNameStrategy();

}
