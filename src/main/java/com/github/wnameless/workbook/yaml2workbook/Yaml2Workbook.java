package com.github.wnameless.workbook.yaml2workbook;

import java.io.StringReader;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.yaml.snakeyaml.LoaderOptions;
import org.yaml.snakeyaml.Yaml;
import org.yaml.snakeyaml.nodes.Node;

public final class Yaml2Workbook {

  public Workbook toWorkbook(String yamlContent) {
    LoaderOptions options = new LoaderOptions();
    options.setProcessComments(true);
    Yaml yaml = new Yaml(options);

    Iterable<Node> roots = yaml.composeAll(new StringReader(yamlContent));

    var workbook = new XSSFWorkbook();
    var sheet = workbook.getSheetAt(0);
    var currentRow = sheet.getRow(0);

    for (var root : roots) {
      traverseNode(root, currentRow);
    }

    return workbook;
  }

  private void traverseNode(Node root, Row currentRow) {
    // Traverse MappingNode, SequenceNode, ScalarNode and print YAML data into workbook
    // Use the settings in Yaml2WorkbookConfig for workbook generation
  }

}
