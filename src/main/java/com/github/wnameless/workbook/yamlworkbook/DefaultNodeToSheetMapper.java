package com.github.wnameless.workbook.yamlworkbook;

import org.yaml.snakeyaml.nodes.Node;

/**
 * Default implementation of {@link NodeToSheetMapper} that maps all nodes to sheet index 0.
 *
 * @author Wei-Ming Wu
 */
public final class DefaultNodeToSheetMapper implements NodeToSheetMapper {

  @Override
  public Integer apply(Node t, Integer u) {
    return 0;
  }

}
