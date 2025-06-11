# .sheet File Format Documentation

## Overview
The `.sheet` file format is a JSON-based format designed to store spreadsheet data that can be efficiently rendered on mobile devices. It supports multiple sheets, formulas, formatting, and cell metadata.

## File Structure
A `.sheet` file contains a JSON array of sheet objects:

```json
[
  {
    "name": "Sheet1",
    "data": [...],
    "config": {...}
  },
  {
    "name": "Sheet2", 
    "data": [...],
    "config": {...}
  }
]
```

## Sheet Object Properties

### `name` (string)
The display name of the sheet. If not provided, defaults to "Sheet1".

### `data` (SheetCell[][])
A 2D array representing the spreadsheet grid, where:
- Outer array represents rows
- Inner arrays represent cells in each row
- Each cell can be a `SheetCell` object or `null`/`undefined` for empty cells

### `config` (object, optional)
Configuration object containing sheet-level settings:
- `merge`: Object defining merged cell ranges
- `rowlen`: Object mapping row indices to custom heights
- `collen`: Object mapping column indices to custom widths

## SheetCell Object Properties

### `v` (string | number, optional)
The display value of the cell. This is what gets shown to the user.

### `f` (string, optional) 
The formula for the cell, starting with "=". When present, this takes precedence over `v` for calculations, but `v` stores the computed result.

### `ct` (object, optional)
Cell type and formatting information:
```json
{
  "fa": "$#,##0.00",  // Format pattern
  "t": "n"            // Type: "n" for number, "s" for string
}
```

### `s` (number, optional)
Style identifier referencing a style definition (implementation dependent).

## Examples

### Basic Sheet
```json
[{
  "name": "Basic Example",
  "data": [
    [{"v": "Name"}, {"v": "Age"}, {"v": "City"}],
    [{"v": "John Doe"}, {"v": 25}, {"v": "New York"}]
  ],
  "config": {
    "merge": {},
    "rowlen": {},
    "collen": {}
  }
}]
```

### Sheet with Formulas
```json
[{
  "name": "Budget Calculator", 
  "data": [
    [{"v": "Item"}, {"v": "Cost"}, {"v": "Quantity"}, {"v": "Total"}],
    [{"v": "Apples"}, {"v": 1.50}, {"v": 10}, {"f": "=B2*C2", "v": 15.00}],
    [{"v": "TOTAL"}, null, null, {"f": "=SUM(D2:D4)", "v": 31.00}]
  ],
  "config": {
    "rowlen": {"0": 30},
    "collen": {"0": 120, "3": 100}
  }
}]
```

### Sheet with Formatting
```json
[{
  "name": "Formatted Data",
  "data": [
    [{"v": "Price", "s": 1}],
    [{"v": 999.99, "ct": {"fa": "$#,##0.00", "t": "n"}}]
  ]
}]
```

## Cell References
Formulas use standard spreadsheet notation:
- `A1`, `B2`, etc. for individual cells
- `A1:B5` for ranges
- Standard functions like `SUM()`, `AVERAGE()`, etc.

## Empty Cells
Empty cells can be represented as:
- `null`
- `undefined` 
- `{}` (empty object)
- Missing array elements

## Configuration Details

### Row Heights (`rowlen`)
Maps row indices (0-based) to heights in pixels:
```json
{
  "rowlen": {
    "0": 30,    // First row is 30px tall
    "4": 35     // Fifth row is 35px tall
  }
}
```

### Column Widths (`collen`)
Maps column indices (0-based) to widths in pixels:
```json
{
  "collen": {
    "0": 120,   // First column is 120px wide
    "1": 80     // Second column is 80px wide
  }
}
```

## Implementation Notes

### Parser Usage
The `SheetParser` class handles parsing and validation:
```typescript
const sheets = SheetParser.parse(fileContent);
const cellValue = SheetParser.getCellValue(cell);
const {rows, cols} = SheetParser.getSheetDimensions(sheetData);
```

### Mobile Optimization
- Compact JSON format minimizes memory usage
- Lazy evaluation of formulas
- Progressive loading for large sheets
- Touch-optimized rendering

### Compatibility
This format is designed to be:
- Forward compatible (new properties ignored by older parsers)
- Backward compatible (missing properties have sensible defaults)  
- Mobile-first (optimized for memory and performance constraints)

## File Extension
Files should use the `.sheet` extension to be recognized by the mobile sheet viewer plugin.