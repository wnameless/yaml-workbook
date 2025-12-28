package com.github.wnameless.workbook.yamlworkbook;

import lombok.Builder;
import lombok.Getter;

/**
 * Configuration for DISPLAY_MODE, controlling how comments are rendered.
 *
 * @author Wei-Ming Wu
 */
@Getter
@Builder
public class DisplayModeConfig {

  public static final DisplayModeConfig DEFAULT = DisplayModeConfig.builder().build();

  // Replaceable types - default to DISPLAY_NAME
  @Builder.Default
  private CommentDisplayOption mappingComment = CommentDisplayOption.DISPLAY_NAME;

  @Builder.Default
  private CommentDisplayOption sequenceComment = CommentDisplayOption.DISPLAY_NAME;

  @Builder.Default
  private CommentDisplayOption keyComment = CommentDisplayOption.DISPLAY_NAME;

  @Builder.Default
  private CommentDisplayOption valueComment = CommentDisplayOption.DISPLAY_NAME;

  // Non-replaceable types - default to HIDDEN for cleaner display
  @Builder.Default
  private CommentVisibility documentComment = CommentVisibility.HIDDEN;

  @Builder.Default
  private CommentVisibility keyValuePairComment = CommentVisibility.HIDDEN;

  @Builder.Default
  private CommentVisibility itemComment = CommentVisibility.HIDDEN;

}
