# Release Notes

## v0.1.0

Initial release with support for:

### Core Features

- **Bidirectional YAML/Excel Conversion**
  - `YamlWorkbookWriter` - Convert YAML to Excel workbooks (.xlsx)
  - `YamlWorkbookReader` - Convert Excel workbooks back to SnakeYAML Node trees
  - `YamlWorkbook` - Static utility class with convenience methods

- **Comment Preservation**
  - Block comments, inline comments, and end comments are preserved during roundtrip conversion
  - Comments stored in appropriate cells with `#` prefix

- **Three Print Modes**
  - `YAML_ORIENTED` - Direct YAML-to-cell mapping (default)
  - `WORKBOOK_READABLE` - Human-readable display with original data in cell comments, configurable via `DisplayModeConfig`
  - `DATA_COLLECT` - Schema-driven data collection with dropdowns from JSON Schema

- **Two Indentation Modes**
  - `CELL_OFFSET` - Uses empty cells for indentation (default)
  - `PREFIX` - Uses prefix markers (1>, 2>, 3>) for deeply nested structures

### JSON Schema Integration

- Generate Excel forms from JSON Schema definitions in `DATA_COLLECT` mode
- `title` property used as display name with original key in cell comment
- `enum` values become dropdown cell validation
- `enumNames` support for human-readable dropdown options
- `DataCollectConfig` options:
  - `highlightRequired` - Highlight required fields with styling
  - `useHiddenSheetsForLongEnums` - Handle dropdowns exceeding 256 character limit
  - `skipAllOf` - Skip allOf merging for conditional schema patterns

### DisplayModeConfig (WORKBOOK_READABLE)

- Fine-grained control over comment rendering in WORKBOOK_READABLE mode
- `CommentDisplayOption` for replaceable types (object, array, key, value comments):
  - `DISPLAY_NAME` - Replace key/value with comment content (default)
  - `HIDDEN` - Show original key/value, ignore comment
  - `COMMENT` - Keep as separate comment cell
- `CommentVisibility` for structural types (document, key-value pair, item comments):
  - `HIDDEN` - Hide comment (default)
  - `COMMENT` - Show comment in separate cell

### Pluggable Strategies

- `WorkbookSyntax` - Customize YAML symbols (frontmatter, comment mark, item mark, etc.)
- `NodeToSheetMapper` - Map YAML documents to specific sheets
- `SheetNameStrategy` - Custom sheet naming conventions
- `IndentPrefixStrategy` - Custom prefix patterns for PREFIX mode

### Multi-Document Support

- Handle multiple YAML documents in a single workbook
- Frontmatter (`---`) used as document separator

### Dependencies

- SnakeYAML 2.x with comment processing enabled
- Apache POI 5.x for Excel generation (XSSFWorkbook)
- Jackson 3.x for JSON processing
- jsonschema-data-generator for JSON Schema integration
