# Tasks: Fix Native Excel Pivot Table

## Phase 1: Infrastructure

### 1.1 Create Template File
- Build a minimal `template-pivot.xlsx` with a native pivot table in LibreOffice
- Template structure: "Tabla dinámica" (pivot display) + "Datos" (source data with headers)
- Place in `public/template-pivot.xlsx`

## Phase 2: Implementation

### 2.1 Reimplement nativePivotExcel.js
- Replace manual XML generation with template loading + data injection
- Fetch template via `fetch()` or read from filesystem (Node for tests, fetch for browser)
- Use SheetJS `json_to_sheet` to generate data sheet XML
- Use JSZip to inject the data sheet into the template
- Update pivot cache `recordCount` and `ref` to match actual data
- Keep `inferNativePivotFields()` logic for field detection

### 2.2 Update Tests
- Update `tests/nativePivotExcel.test.js` to test the new template approach
- Verify output contains valid pivot table parts

## Phase 3: Verification

### 3.1 Build Verification
- Run `npm run build` successfully
- Verify bundle includes the new code

### 3.2 Runtime Verification
- Generate test file locally
- Open in LibreOffice and confirm values are correct
- Deliver to user for Excel testing
