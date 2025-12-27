package com.github.wnameless.workbook.yamlworkbook;

/**
 * Visibility options for non-replaceable comment types (DOCUMENT, KEY_VALUE_PAIR, ITEM).
 *
 * @author Wei-Ming Wu
 */
public enum CommentVisibility {

  /** Hide comment, show only structural element */
  HIDDEN,

  /** Show comment (current behavior) */
  COMMENT

}
