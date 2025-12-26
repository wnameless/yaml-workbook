package com.github.wnameless.workbook.yamlworkbook;

public enum PrintMode {

  /** Direct YAML-to-cell mapping, no transformation */
  YAML_ORIENTED,

  /** Human-readable with original data preserved in cell comments for roundtrip support */
  WORKBOOK_READABLE,

  /** Schema-driven data collection with dropdowns and metadata from JSON Schema */
  DATA_COLLECT

}
