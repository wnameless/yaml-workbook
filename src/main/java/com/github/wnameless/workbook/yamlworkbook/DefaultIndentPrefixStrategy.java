package com.github.wnameless.workbook.yamlworkbook;

/**
 * Default implementation of {@link IndentPrefixStrategy} using numeric prefixes.
 * <p>
 * Pattern: "1>", "2>", "3>", etc.
 * <ul>
 * <li>Level 0: empty string (root level, no prefix)</li>
 * <li>Level 1: "1>"</li>
 * <li>Level 2: "2>"</li>
 * <li>Level N: "N>"</li>
 * </ul>
 *
 * @author Wei-Ming Wu
 */
public class DefaultIndentPrefixStrategy implements IndentPrefixStrategy {

  private static final String SUFFIX = ">";

  @Override
  public String generatePrefix(int indentLevel) {
    if (indentLevel <= 0) {
      return "";
    }
    return indentLevel + SUFFIX;
  }

  @Override
  public int parsePrefix(String prefix) {
    if (prefix == null || prefix.isEmpty()) {
      return 0;
    }
    if (!prefix.endsWith(SUFFIX)) {
      return -1;
    }
    try {
      String numPart = prefix.substring(0, prefix.length() - SUFFIX.length());
      int level = Integer.parseInt(numPart);
      return level > 0 ? level : -1;
    } catch (NumberFormatException e) {
      return -1;
    }
  }

}
