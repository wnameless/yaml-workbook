package com.github.wnameless.workbook.yamlworkbook;

import org.yaml.snakeyaml.nodes.Node;

public final class DefaultYamlNodeClassifier implements YamlNodeSheetClassifier {

  @Override
  public Integer apply(Node t, Integer u) {
    return 0;
  }

}
