# Native Pivot Excel Specification

## Purpose

Define the requirements for generating native Excel pivot tables that render correctly in both Microsoft Excel and LibreOffice Calc.

## Requirements

### Requirement: Template-Based Pivot Generation

The system MUST generate native Excel pivot tables using a pre-built template file rather than hand-crafted OpenXML.

#### Scenario: Export actuaciones dataset

- GIVEN a dataset with columns `Año`, `Mes`, and `Nombre de la actuación`
- WHEN the user exports a native pivot table
- THEN the downloaded `.xlsx` MUST contain a native pivot table with `Año` and `Nombre de la actuación` as row fields
- AND `Mes` as column field
- AND `Cuenta de Nombre de la actuación` as the value aggregation
- AND all cell values MUST be non-zero when source data has matching records

#### Scenario: Export generic dataset

- GIVEN a dataset with numeric and text columns
- WHEN the user exports a native pivot table
- THEN the system MUST infer appropriate row, column, and value fields
- AND the pivot table MUST aggregate numeric fields with SUM and text fields with COUNT

### Requirement: Excel Compatibility

The generated `.xlsx` MUST open correctly in Microsoft Excel without showing zero values.

#### Scenario: Open in Microsoft Excel

- GIVEN a generated pivot table `.xlsx`
- WHEN opened in Microsoft Excel 2016 or later
- THEN the pivot table MUST display correct aggregated values
- AND the pivot table MUST be editable (drillable) by the user

#### Scenario: Open in LibreOffice Calc

- GIVEN a generated pivot table `.xlsx`
- WHEN opened in LibreOffice Calc
- THEN the pivot table MUST display the same values as in Excel

### Requirement: Data Source Integrity

The data source sheet injected into the template MUST preserve data types (numbers as numbers, strings as strings).

#### Scenario: Numeric values in source data

- GIVEN source data with numeric columns
- WHEN injected into the template
- THEN the resulting sheet cells MUST have numeric type, not text
- AND the pivot table MUST be able to SUM those values

## MODIFIED Requirements

None.

## REMOVED Requirements

None.
