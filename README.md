![Maven Central Version](https://img.shields.io/maven-central/v/com.github.wnameless.workbook/yaml-workbook)
[![codecov](https://codecov.io/gh/wnameless/yaml-workbook/graph/badge.svg)](https://codecov.io/gh/wnameless/yaml-workbook)

yaml-workbook
=============
A Java library for bidirectional conversion between YAML and Excel workbooks. It uses SnakeYAML for YAML parsing and Apache POI for Excel generation, supporting roundtrip conversion with comment preservation.

## Key Features

- **Bidirectional Conversion** — Convert YAML to Excel and Excel back to YAML with full roundtrip support
- **Comment Preservation** — Block, inline, and end comments are preserved during conversion
- **Multiple Output Modes** — Three output modes for different use cases (YAML-oriented, human-readable, data collection)
- **JSON Schema Integration** — Generate Excel forms with dropdowns from JSON Schema definitions
- **Flexible Indentation** — Choose between cell-offset or prefix-based indentation for deep nesting
- **Pluggable Strategies** — Customize sheet naming, node mapping, syntax symbols, and more
- **Multi-Document Support** — Handle multiple YAML documents in a single workbook

## Purpose
Converts a YAML document
```yaml
# User profile
name: John Doe  # display name
age: 30
hobbies:
  - reading
  - coding
```
into an Excel workbook

| A | B | C |
|---|---|---|
| --- | | |
| # User profile | | |
| name | John Doe | # display name |
| age | 30 | |
| hobbies | | |
| - | reading | |
| - | coding | |

and back to YAML with comments preserved.

# Maven Repo
```xml
<dependency>
  <groupId>com.github.wnameless.workbook</groupId>
  <artifactId>yaml-workbook</artifactId>
  <version>${newestVersion}</version>
  <!-- Newest version shows in the maven-central badge above -->
</dependency>
```

# Quick Start

## YAML to Excel
```java
String yaml = """
    name: John Doe
    age: 30
    email: john@example.com
    """;

// Simple conversion
Workbook workbook = YamlWorkbook.toWorkbook(yaml);

// Save to file
try (FileOutputStream fos = new FileOutputStream("output.xlsx")) {
    workbook.write(fos);
}
```

## Excel to YAML
```java
// Read workbook
Workbook workbook = WorkbookFactory.create(new File("output.xlsx"));

// Convert to YAML string
String yaml = YamlWorkbook.toYaml(workbook);

// Or get SnakeYAML Node objects for further processing
List<Node> nodes = YamlWorkbook.fromWorkbook(workbook);
```

## Roundtrip Conversion
```java
// Original YAML with comments
String original = """
    # Configuration file
    server:
      host: localhost  # server hostname
      port: 8080
    """;

// Convert to workbook and back
Workbook workbook = YamlWorkbook.toWorkbook(original);
String restored = YamlWorkbook.toYaml(workbook);
// Comments are preserved!
```

# Output Modes

The library provides three output modes for different use cases:

```java
// YAML_ORIENTED (default) - Direct YAML-to-cell mapping
Workbook wb1 = YamlWorkbook.writerBuilder()
    .outputMode(OutputMode.YAML_ORIENTED)
    .build()
    .toWorkbook(yamlReader);

// DISPLAY_MODE - Human-readable with original data in cell comments
Workbook wb2 = YamlWorkbook.writerBuilder()
    .outputMode(OutputMode.DISPLAY_MODE)
    .build()
    .toWorkbook(yamlReader);

// FORM_MODE - Schema-driven forms with dropdowns from JSON Schema
Workbook wb3 = YamlWorkbook.writerBuilder()
    .outputMode(OutputMode.FORM_MODE)
    .jsonSchema(jsonSchemaString)
    .build()
    .toWorkbook();
```

| Mode | Description | Use Case |
|------|-------------|----------|
| `YAML_ORIENTED` | Direct mapping, comments as cells | Development, debugging |
| `DISPLAY_MODE` | Display names from comments, originals in cell comments | End-user documentation |
| `FORM_MODE` | JSON Schema-driven with dropdowns | Data entry forms |

# Indentation Modes

Choose how nested structures are represented:

```java
// CELL_OFFSET (default) - Uses empty cells for indentation
// Cell index = indentLevel * cellNum
Workbook wb1 = YamlWorkbook.writerBuilder()
    .indentationMode(IndentationMode.CELL_OFFSET)
    .build()
    .toWorkbook(yamlReader);

// PREFIX - Uses prefix markers (1>, 2>, 3>) in first cell
// Better for deeply nested structures
Workbook wb2 = YamlWorkbook.prefixWriterBuilder()
    .build()
    .toWorkbook(yamlReader);
```

**CELL_OFFSET example:**
| A | B | C | D |
|---|---|---|---|
| name | John | | |
| address | | | |
| | city | NYC | |
| | zip | 10001 | |

**PREFIX example:**
| A | B | C |
|---|---|---|
| name | John | |
| address | | |
| 1> | city | NYC |
| 1> | zip | 10001 |

# JSON Schema Integration (FORM_MODE Mode)

Generate Excel forms with dropdowns and validation from JSON Schema:

```java
String jsonSchema = """
    {
      "type": "object",
      "properties": {
        "status": {
          "type": "string",
          "title": "Status",
          "enum": ["active", "inactive", "pending"],
          "enumNames": ["Active", "Inactive", "Pending"]
        },
        "priority": {
          "type": "integer",
          "title": "Priority Level",
          "enum": [1, 2, 3]
        }
      }
    }
    """;

Workbook workbook = YamlWorkbook.writerBuilder()
    .outputMode(OutputMode.FORM_MODE)
    .jsonSchema(jsonSchema)
    .formModeConfig(FormModeConfig.builder()
        .useHiddenSheetsForLongEnums(true)  // Handle large dropdowns
        .skipAllOf(true)                     // Skip allOf for conditional schemas
        .build())
    .build()
    .toWorkbook();
```

Features:
- `title` property becomes display name, original key stored in cell comment
- `enum` values become dropdown cell validation
- `enumNames` (when present) become dropdown display values

## FormModeConfig Options

| Option | Type | Default | Description |
|--------|------|---------|-------------|
| `highlightRequired` | boolean | false | Highlight required fields with styling |
| `useHiddenSheetsForLongEnums` | boolean | false | Use hidden sheets for dropdowns exceeding 256 chars |
| `skipAllOf` | boolean | false | Skip allOf merging for conditional schema patterns |

# DisplayModeConfig (DISPLAY_MODE)

Customize how comments are rendered in DISPLAY_MODE mode:

```java
Workbook workbook = YamlWorkbook.writerBuilder()
    .outputMode(OutputMode.DISPLAY_MODE)
    .displayModeConfig(DisplayModeConfig.builder()
        .mappingComment(CommentDisplayOption.DISPLAY_NAME)  // Use comment as display name for mappings
        .sequenceComment(CommentDisplayOption.COMMENT)  // Keep as separate comment cell for sequences
        .documentComment(CommentVisibility.HIDDEN)      // Hide document-level comments
        .build())
    .build()
    .toWorkbook(yamlReader);
```

## DisplayModeConfig Options

| Option | Type | Default | Description |
|--------|------|---------|-------------|
| `mappingComment` | CommentDisplayOption | DISPLAY_NAME | How to handle comments before mappings |
| `sequenceComment` | CommentDisplayOption | DISPLAY_NAME | How to handle comments before sequences |
| `keyComment` | CommentDisplayOption | DISPLAY_NAME | How to handle key inline comments |
| `valueComment` | CommentDisplayOption | DISPLAY_NAME | How to handle value inline comments |
| `documentComment` | CommentVisibility | HIDDEN | How to handle document-level comments |
| `keyValuePairComment` | CommentVisibility | HIDDEN | How to handle key-value pair block comments |
| `itemComment` | CommentVisibility | HIDDEN | How to handle sequence item block comments |

## CommentDisplayOption

| Option | Description |
|--------|-------------|
| `DISPLAY_NAME` | Replace key/value with comment content (default) |
| `HIDDEN` | Show original key/value, ignore comment |
| `COMMENT` | Keep as separate comment cell |

## CommentVisibility

| Option | Description |
|--------|-------------|
| `HIDDEN` | Hide comment (default) |
| `COMMENT` | Show comment in separate cell |

# Customization

## Custom Workbook Syntax
```java
// Customize YAML symbols
WorkbookSyntax customSyntax = new WorkbookSyntax() {
    public String getFrontmatter() { return "---"; }
    public String getCommentMark() { return "#"; }
    public String getValueEscapeMark() { return "\\"; }
    public String getItemMark() { return "-"; }
    public Short getIndentationCellNum() { return 2; }  // 2 cells per indent
};

Workbook workbook = YamlWorkbook.writerBuilder()
    .workbookSyntax(customSyntax)
    .build()
    .toWorkbook(yamlReader);
```

## Custom Sheet Naming
```java
// Custom sheet names
SheetNameStrategy customNaming = index -> "Data_" + (index + 1);

Workbook workbook = YamlWorkbook.writerBuilder()
    .sheetNameStrategy(customNaming)
    .build()
    .toWorkbook(yamlReader);
```

## Multi-Document to Multiple Sheets
```java
// Map each YAML document to a separate sheet
NodeToSheetMapper oneDocPerSheet = (node, index) -> index;

Workbook workbook = YamlWorkbook.writerBuilder()
    .nodeToSheetMapper(oneDocPerSheet)
    .build()
    .toWorkbook(yamlReader);
```

## Custom Indent Prefix
```java
// Use custom prefix pattern (e.g., ">>", ">>>>", ">>>>>>")
IndentPrefixStrategy customPrefix = new IndentPrefixStrategy() {
    public String generatePrefix(int level) {
        return ">>".repeat(level);
    }
    public int parsePrefix(String prefix) {
        if (prefix == null || !prefix.matches("(>>)+")) return -1;
        return prefix.length() / 2;
    }
};

Workbook workbook = YamlWorkbook.writerBuilder()
    .indentationMode(IndentationMode.PREFIX)
    .indentPrefixStrategy(customPrefix)
    .build()
    .toWorkbook(yamlReader);
```

# Configuration Reference

## Writer Configuration

| Option | Type | Default | Description |
|--------|------|---------|-------------|
| `outputMode` | OutputMode | YAML_ORIENTED | Output format mode |
| `indentationMode` | IndentationMode | CELL_OFFSET | Indentation representation |
| `workbookSyntax` | WorkbookSyntax | DEFAULT | YAML symbols configuration |
| `nodeToSheetMapper` | NodeToSheetMapper | DEFAULT | Document-to-sheet mapping |
| `sheetNameStrategy` | SheetNameStrategy | DEFAULT | Sheet naming convention |
| `indentPrefixStrategy` | IndentPrefixStrategy | DEFAULT | Prefix generation (for PREFIX mode) |
| `displayModeConfig` | DisplayModeConfig | DEFAULT | DISPLAY_MODE mode options |
| `formModeConfig` | FormModeConfig | DEFAULT | FORM_MODE mode options |
| `jsonSchema` | String | null | JSON Schema for FORM_MODE mode |

## Reader Configuration

| Option | Type | Default | Description |
|--------|------|---------|-------------|
| `outputMode` | OutputMode | YAML_ORIENTED | Expected workbook format |
| `indentationMode` | IndentationMode | CELL_OFFSET | Expected indentation style |
| `workbookSyntax` | WorkbookSyntax | DEFAULT | YAML symbols configuration |
| `sheetNameStrategy` | SheetNameStrategy | DEFAULT | Sheet naming convention |
| `indentPrefixStrategy` | IndentPrefixStrategy | DEFAULT | Prefix parsing (for PREFIX mode) |

# API Overview

## YamlWorkbook (Utility Class)
```java
// Builder access
YamlWorkbook.writerBuilder()        // Standard writer builder
YamlWorkbook.readerBuilder()        // Standard reader builder
YamlWorkbook.prefixWriterBuilder()  // PREFIX mode writer
YamlWorkbook.prefixReaderBuilder()  // PREFIX mode reader

// Convenience methods
YamlWorkbook.toWorkbook(String yaml)              // YAML string to workbook
YamlWorkbook.toWorkbook(InputStream is)           // InputStream to workbook
YamlWorkbook.fromWorkbook(Workbook wb)            // Workbook to Node list
YamlWorkbook.toYaml(Workbook wb)                  // Workbook to YAML string
```

## YamlWorkbookWriter
```java
YamlWorkbookWriter writer = YamlWorkbook.writerBuilder()
    .outputMode(OutputMode.YAML_ORIENTED)
    .build();

// From Reader
Workbook wb = writer.toWorkbook(new StringReader(yaml));

// From JSON Schema (FORM_MODE mode only)
Workbook wb = writer.toWorkbook();
```

## YamlWorkbookReader
```java
YamlWorkbookReader reader = YamlWorkbook.readerBuilder()
    .outputMode(OutputMode.YAML_ORIENTED)
    .build();

List<Node> nodes = reader.fromWorkbook(workbook);
```

# Requirements

- Java 17 or higher

# Dependencies

- SnakeYAML 2.x for YAML parsing (with comment support)
- Apache POI 5.x for Excel workbook generation (XSSFWorkbook for .xlsx)
- Jackson 3.x for JSON processing (via jsonschema-data-generator)
- jsonschema-data-generator for JSON Schema integration in FORM_MODE mode
- Lombok for builder pattern support

# License

Apache License 2.0
