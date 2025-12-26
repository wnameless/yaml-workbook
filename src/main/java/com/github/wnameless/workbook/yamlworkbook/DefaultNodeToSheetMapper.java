package com.github.wnameless.workbook.yamlworkbook;

import org.yaml.snakeyaml.nodes.Node;

public final class DefaultNodeToSheetMapper implements NodeToSheetMapper {

  @Override
  public Integer apply(Node t, Integer u) {
    return 0;
  }

}
