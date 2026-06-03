# Proposal: Fix Native Excel Pivot Table Generation

## Intent

The current implementation generates native Excel pivot tables by manually constructing OpenXML, which is fragile and produces files that work in LibreOffice but show zero values in Microsoft Excel. We need a robust, maintainable approach that produces correct pivot tables in both Excel and LibreOffice.

## Scope

### In Scope
- Replace manual OpenXML generation with a template-based approach
- Create a pre-built `.xlsx` template containing a valid native pivot table
- Inject user data into the template's source sheet while preserving the pivot table definition
- Support the actuaciones dataset structure: Año + Mes + Nombre de la actuación (COUNT)
- Update tests to verify the new approach

### Out of Scope
- Adding new pivot table layouts or configurations beyond the current use case
- Supporting Excel versions older than 2016
- Modifying the UI/UX of the pivot table export button

## Capabilities

### New Capabilities
- `native-pivot-excel`: Robust generation of native Excel pivot tables using pre-built templates

### Modified Capabilities
- None (this is a pure implementation fix, spec behavior remains the same)

## Approach

Use a **template-based approach**:
1. Create a `template-pivot.xlsx` file with a pre-configured native pivot table (built in LibreOffice/Excel)
2. At export time, load the template via JSZip
3. Replace the `xl/worksheets/sheet2.xml` (data sheet) with user data using SheetJS
4. Keep the pivot table definition, cache, and relationships intact
5. Return the modified workbook as a Blob for download

This guarantees 100% Excel compatibility because the pivot table XML is created by Excel/LibreOffice itself, not by hand.

## Affected Areas

| Area | Impact | Description |
|------|--------|-------------|
| `src/utils/nativePivotExcel.js` | Modified | Replace manual XML with template injection |
| `public/template-pivot.xlsx` | New | Pre-built template with native pivot table |
| `tests/nativePivotExcel.test.js` | Modified | Update tests for template approach |

## Risks

| Risk | Likelihood | Mitigation |
|------|------------|------------|
| Template file becomes large | Low | Keep template minimal (empty data + pivot structure) |
| Template breaks with future Excel versions | Low | Template uses standard OpenXML features |
| Data injection misaligns headers | Med | Validate header count matches template before injection |

## Rollback Plan

Revert `nativePivotExcel.js` to the previous commit and remove the template file. The old manual-XML approach is preserved in git history.

## Dependencies

- JSZip (already installed)
- xlsx (SheetJS) (already installed)

## Success Criteria

- [ ] Generated `.xlsx` opens correctly in Microsoft Excel with non-zero values
- [ ] Generated `.xlsx` opens correctly in LibreOffice Calc
- [ ] Test passes verifying pivot table parts exist in the output
- [ ] Build completes without errors
