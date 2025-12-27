package com.github.wnameless.workbook.yamlworkbook;

import lombok.Builder;
import lombok.Getter;

/**
 * Configuration for DATA_COLLECT mode.
 * <p>
 * Most behaviors are mandatory and not configurable:
 * <ul>
 * <li>{@code title} in schema is used as display name, original key stored in cell comment</li>
 * <li>{@code enum} creates dropdown cell validation</li>
 * <li>{@code enumNames} (when present) are used as dropdown display values, actual enum values
 * stored in cell comment</li>
 * </ul>
 *
 * @author Wei-Ming Wu
 */
@Getter
@Builder
public class DataCollectConfig {

  public static final DataCollectConfig DEFAULT = DataCollectConfig.builder().build();

  /** Optional: highlight required fields with styling */
  @Builder.Default
  private boolean highlightRequired = false;

  /**
   * When true, enum dropdowns exceeding 256 characters will write values to a hidden sheet and use
   * named ranges (bypasses POI's 256 char limit). When false, dropdowns are truncated with a
   * warning.
   */
  @Builder.Default
  private boolean useHiddenSheetsForLongEnums = false;

  /**
   * When true, skips allOf merging during JSON Schema processing. Use this when schemas use allOf
   * for conditional statements (if/then/else) that are not supported by the schema skeleton
   * generator.
   */
  @Builder.Default
  private boolean skipAllOf = false;

}
