package com.github.wnameless.workbook.yamlworkbook;

/**
 * Strategy for generating and parsing indent prefixes in PREFIX indentation mode.
 * <p>
 * Used by {@link YamlWorkbookWriter} to generate prefix strings for each indent level, and by
 * {@link YamlWorkbookReader} to parse prefix strings back to indent levels.
 */
public interface IndentPrefixStrategy {

  /**
   * Default implementation using numeric prefixes: "1>", "2>", "3>", etc.
   */
  public static final IndentPrefixStrategy DEFAULT = new DefaultIndentPrefixStrategy();

  /**
   * Generates a prefix string for the given indent level.
   *
   * @param indentLevel the indent level (0 = root, 1 = first nest, etc.)
   * @return the prefix string, or empty string for level 0
   */
  String generatePrefix(int indentLevel);

  /**
   * Parses a prefix string to determine the indent level.
   *
   * @param prefix the prefix string from the cell
   * @return the indent level (>= 0), or -1 if not a valid prefix
   */
  int parsePrefix(String prefix);

  /**
   * Checks if a string is a valid prefix.
   *
   * @param value the cell value to check
   * @return true if this is a valid prefix (parsePrefix returns >= 0)
   */
  default boolean isPrefix(String value) {
    return parsePrefix(value) >= 0;
  }

}
