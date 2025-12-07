package com.github.wnameless.workbook.yamlworkbook;

import java.util.function.BiFunction;
import org.yaml.snakeyaml.nodes.Node;

public interface NodeToSheetMapper extends BiFunction<Node, Integer, Integer> {

  public static final NodeToSheetMapper DEFAULT = new DefaultNodeToSheetMapper();

  Integer apply(Node node, Integer nodeIndex);

}
