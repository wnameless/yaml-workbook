package com.github.wnameless.workbook.yamlworkbook;

import java.io.StringReader;
import org.apache.poi.ss.usermodel.Workbook;
import lombok.experimental.UtilityClass;

@UtilityClass
public class YamlWorkbook {

  public YamlWorkbookWriter.YamlWorkbookWriterBuilder writerBuilder() {
    return YamlWorkbookWriter.builder();
  }

  public Workbook toWorkbook(String yamlContent, String... yamlContents) {
    StringReader[] yamlContentReaders = new StringReader[yamlContents.length];
    for (int i = 0; i < yamlContents.length; i++) {
      yamlContentReaders[i] = new StringReader(yamlContents[i]);
    }
    return YamlWorkbookWriter.builder().build().toWorkbook(new StringReader(yamlContent),
        yamlContentReaders);
  }

  public Workbook toWorkbook(StringReader yamlContent, StringReader... yamlContents) {
    return YamlWorkbookWriter.builder().build().toWorkbook(yamlContent, yamlContents);
  }

}
