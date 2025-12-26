package com.github.wnameless.workbook.yamlworkbook;

import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import org.yaml.snakeyaml.DumperOptions.FlowStyle;
import org.yaml.snakeyaml.DumperOptions.ScalarStyle;
import org.yaml.snakeyaml.nodes.MappingNode;
import org.yaml.snakeyaml.nodes.Node;
import org.yaml.snakeyaml.nodes.NodeTuple;
import org.yaml.snakeyaml.nodes.ScalarNode;
import org.yaml.snakeyaml.nodes.SequenceNode;
import org.yaml.snakeyaml.nodes.Tag;
import com.github.wnameless.json.jsonschemadatagenerator.ObjectMapperFactory;
import lombok.experimental.UtilityClass;
import tools.jackson.databind.JsonNode;
import tools.jackson.databind.node.ArrayNode;
import tools.jackson.databind.node.BooleanNode;
import tools.jackson.databind.node.NullNode;
import tools.jackson.databind.node.NumericNode;
import tools.jackson.databind.node.ObjectNode;
import tools.jackson.databind.node.StringNode;

/**
 * Utility class to convert Jackson JsonNode to SnakeYAML Node.
 */
@UtilityClass
public class JsonNodeToYamlNodeConverter {

  /**
   * Converts a Jackson JsonNode to a SnakeYAML Node.
   *
   * @param jsonNode the Jackson JsonNode to convert
   * @return the equivalent SnakeYAML Node
   */
  public static Node convert(JsonNode jsonNode) {
    if (jsonNode == null || jsonNode instanceof NullNode) {
      return new ScalarNode(Tag.NULL, "null", null, null, ScalarStyle.PLAIN);
    }

    if (jsonNode instanceof ObjectNode) {
      return convertObject(jsonNode);
    } else if (jsonNode instanceof ArrayNode) {
      return convertArray(jsonNode);
    } else {
      return convertScalar(jsonNode);
    }
  }

  private static MappingNode convertObject(JsonNode objectNode) {
    List<NodeTuple> tuples = new ArrayList<>();

    Iterator<Map.Entry<String, JsonNode>> fields = objectNode.properties().iterator();
    while (fields.hasNext()) {
      Map.Entry<String, JsonNode> entry = fields.next();
      ScalarNode keyNode = new ScalarNode(Tag.STR, entry.getKey(), null, null, ScalarStyle.PLAIN);
      Node valueNode = convert(entry.getValue());
      tuples.add(new NodeTuple(keyNode, valueNode));
    }

    return new MappingNode(Tag.MAP, tuples, FlowStyle.BLOCK);
  }

  private static SequenceNode convertArray(JsonNode arrayNode) {
    List<Node> items = new ArrayList<>();

    for (JsonNode element : arrayNode) {
      items.add(convert(element));
    }

    return new SequenceNode(Tag.SEQ, items, FlowStyle.BLOCK);
  }

  private static ScalarNode convertScalar(JsonNode scalarNode) {
    if (scalarNode instanceof StringNode) {
      return new ScalarNode(Tag.STR, scalarNode.asString(), null, null, ScalarStyle.PLAIN);
    } else if (scalarNode instanceof NumericNode) {
      if (scalarNode.isIntegralNumber()) {
        return new ScalarNode(Tag.INT, scalarNode.asString(), null, null, ScalarStyle.PLAIN);
      } else {
        return new ScalarNode(Tag.FLOAT, scalarNode.asString(), null, null, ScalarStyle.PLAIN);
      }
    } else if (scalarNode instanceof BooleanNode) {
      return new ScalarNode(Tag.BOOL, scalarNode.asString(), null, null, ScalarStyle.PLAIN);
    } else if (scalarNode instanceof NullNode) {
      return new ScalarNode(Tag.NULL, "null", null, null, ScalarStyle.PLAIN);
    } else {
      // Fallback: use string representation
      return new ScalarNode(Tag.STR, scalarNode.asString(), null, null, ScalarStyle.PLAIN);
    }
  }

  /**
   * Converts a JsonNode to a Map for alternative processing.
   *
   * @param jsonNode the Jackson JsonNode to convert
   * @return a Map representation of the JsonNode
   */
  @SuppressWarnings("unchecked")
  public static Map<String, Object> toMap(JsonNode jsonNode) {
    return ObjectMapperFactory.getObjectMapper().convertValue(jsonNode, Map.class);
  }

}
