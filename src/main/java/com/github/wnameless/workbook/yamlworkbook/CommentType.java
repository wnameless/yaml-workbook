package com.github.wnameless.workbook.yamlworkbook;

/**
 * Types of comments based on their position in YAML structure.
 *
 * @author Wei-Ming Wu
 */
public enum CommentType {

  /** Before frontmatter (---) */
  DOCUMENT,

  /** Before mapping node */
  OBJECT,

  /** Before sequence node */
  ARRAY,

  /** Before/inline with key */
  KEY,

  /** Inline with value */
  VALUE,

  /** Describes whole key-value tuple */
  KEY_VALUE_PAIR,

  /** Before/inline with sequence item */
  ITEM

}
