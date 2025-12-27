package com.github.wnameless.workbook.yamlworkbook;

/**
 * Display options for replaceable comment types (OBJECT, ARRAY, KEY, VALUE).
 *
 * @author Wei-Ming Wu
 */
public enum CommentDisplayOption {

  /** Replace key/value with comment content */
  DISPLAY_NAME,

  /** Show original key/value, ignore comment */
  HIDDEN,

  /** Keep as comment cell (current behavior) */
  COMMENT

}
