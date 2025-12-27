package com.github.wnameless.workbook.yamlworkbook;

import java.util.function.BiFunction;
import org.yaml.snakeyaml.nodes.Node;

/**
 * Strategy for mapping YAML nodes to workbook sheet indices.
 * <p>
 * Determines which sheet a YAML node should be written to based on the node content and its
 * position in the document sequence.
 *
 * @author Wei-Ming Wu
 */
public interface NodeToSheetMapper extends BiFunction<Node, Integer, Integer> {

  /** Default implementation that maps all nodes to sheet index 0. */
  public static final NodeToSheetMapper DEFAULT = new DefaultNodeToSheetMapper();

  /**
   * Determines the sheet index for a given YAML node.
   *
   * @param node the YAML node to map
   * @param nodeIndex the index of this node in the document sequence
   * @return the target sheet index (0-based)
   */
  Integer apply(Node node, Integer nodeIndex);

}

