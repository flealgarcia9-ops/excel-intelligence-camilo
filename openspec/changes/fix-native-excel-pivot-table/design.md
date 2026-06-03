# Design: Fix Native Excel Pivot Table

## Architecture Decision

**Decision**: Replace hand-crafted OpenXML with a template-based approach.

**Rationale**: Manual OpenXML generation for pivot tables is extremely complex (pivotCacheDefinition, pivotCacheRecords, pivotTableDefinition, relationships, content types). Minor deviations cause Excel to reject the pivot table (showing zeros). A pre-built template guarantees 100% valid OpenXML because Excel/LibreOffice created it.

## Approach

1. **Create template**: Build a minimal `.xlsx` with a native pivot table in LibreOffice, using placeholder data.
2. **Store template**: Place `template-pivot.xlsx` in `public/` so Vite includes it in the build.
3. **Inject data at runtime**:
   - Fetch the template as an ArrayBuffer
   - Load it with JSZip
   - Use SheetJS to generate a new data sheet from user data
   - Replace the template's data sheet with the generated one
   - Update the `worksheetSource` reference in the pivot cache if needed
   - Generate the final `.xlsx`

## Data Flow

```
User Data (JSON array)
    ↓
inferNativePivotFields() → determines row/col/value fields
    ↓
SheetJS writes data sheet as XML string
    ↓
JSZip loads template-pivot.xlsx
    ↓
Replace xl/worksheets/sheet2.xml with data sheet XML
    ↓
Update pivot cache source reference to match data range
    ↓
Generate final .xlsx → download
```

## Key Technical Details

- Template has 2 sheets: "Tabla dinámica" (pivot display) and "Datos" (source data)
- The pivot table in the template references "Datos" as its source
- When we replace "Datos", we must also update the `recordCount` and `ref` in the pivot cache
- SheetJS `json_to_sheet` generates valid sheet XML that we inject directly

## Rollback

If the template approach fails, revert to the previous commit of `nativePivotExcel.js`.
