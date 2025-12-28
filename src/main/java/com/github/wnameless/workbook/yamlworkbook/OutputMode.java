package com.github.wnameless.workbook.yamlworkbook;

/**
 * Defines the output format modes for YAML to workbook conversion.
 *
 * @author Wei-Ming Wu
 */
public enum OutputMode {

  /** Direct YAML-to-cell mapping, no transformation */
  YAML_ORIENTED,

  /** Human-readable with original data preserved in cell comments for roundtrip support */
  DISPLAY_MODE,

  /** Schema-driven data collection with dropdowns and metadata from JSON Schema */
  FORM_MODE

}
