/**
 * Module for bidirectional conversion between YAML and Excel workbooks.
 */
module com.github.wnameless.workbook.yamlworkbook {

  // Required modules
  requires transitive org.yaml.snakeyaml;
  requires transitive org.apache.poi.ooxml;
  requires transitive tools.jackson.databind;
  requires com.github.wnameless.json.jsonschemadatagenerator;
  requires java.logging;
  requires static lombok;

  // Export public API
  exports com.github.wnameless.workbook.yamlworkbook;

}
