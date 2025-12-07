package com.github.wnameless.workbook.yamlworkbook;

public enum PrintMode {

  /** Direct YAML-to-cell mapping, no transformation */
  YAML_ORIENTED,

  /** Human-readable, display only */
  WORKBOOK_DISPLAY,

  /** Human-readable with roundtrip support (future) */
  WORKBOOK_ROUNDTRIP

}
