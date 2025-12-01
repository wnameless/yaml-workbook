package com.github.wnameless.workbook.yamlworkbook;

import java.util.function.BiFunction;
import org.yaml.snakeyaml.nodes.Node;

public interface YamlNodeSheetClassifier extends BiFunction<Node, Integer, Integer> {

  public static final YamlNodeSheetClassifier DEFAULT = new DefaultYamlNodeClassifier();

  Integer apply(Node node, Integer nodeIndex);

}
