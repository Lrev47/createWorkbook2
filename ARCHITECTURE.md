# Architecture Overview

## System Purpose

VBA-based workbook generator for fleet asset transactions. Creates pre-formatted Excel workbooks for tracking equipment usage, returns, and swaps across the organization.

## Transaction Flow

```
validate → copy template → map data → save → generate SharePoint link
```

1. **Validate**: Form fields checked for required data and business rules
2. **Copy Template**: Appropriate template selected from 10 variants
3. **Map Data**: Form values written to template cells
4. **Save**: Workbook saved to year/customer folder structure
5. **Generate Link**: SharePoint URL created via OneDrive registry lookup

## Module Responsibilities

### EntryPoint.bas
Main UI controller and orchestrator.
- `Main()` - Sets up workbook UI: dark navy background, header, dropdown, Submit button
- `ProcessTransaction()` - Orchestrates the full transaction flow
- `InstallChangeHandler()` - Injects Worksheet_Change event into Sheet1
- `BuildInfoBox()` - Dynamic info panel that adjusts per order type

### TX_NewUsage.bas
New Usage transaction form (most complex).
- `BuildForm()` - 40+ row form layout with conditional sections
- `PopulateData()` - Maps form values to template cells
- `GetTemplate()` - Selects from 8 template variants based on 3 boolean flags
- Handles Kehe customer detection, Stock Equipment flag, Has Return flag

### TX_Return.bas
Return transaction form with bulk entry.
- `BuildForm()` - Form with 300-row serial number entry area
- `PopulateData()` - Handles CRDB lookups for serial data
- Conditional formatting for the bulk entry rows

### TX_Swap.bas
Swap transaction form.
- `BuildForm()` - Form with Dealer ID entry field
- `PopulateData()` - Equipment type logic (Forklift vs Scrubber variants)
- Simpler layout than New Usage

### Dispatcher.bas
Routes order type changes to transaction modules.
- `OnOrderTypeChange()` - Triggered by dropdown change
- `ClearFormArea()` - Resets form between order type switches
- Maps dropdown values to TX module names

### PathHelper.bas
Path operations with security focus.
- `SanitizeFilename()` - Removes invalid characters, trims spaces
- `EnsureFolderExists()` - Creates folder hierarchy recursively
- `BuildOutputPath()` - Constructs year/customer/filename structure

### FileHelper.bas
Template and file operations.
- `ValidateTemplate()` - Checks template exists and is accessible
- `CopyTemplate()` - Creates workbook copy for editing
- `HandleExistingFile()` - User prompt when file already exists

### SharePointHelper.bas
OneDrive sync to SharePoint URL conversion.
- `GetSharePointUrl()` - Main entry point
- Registry lookup for OneDrive Business sync providers
- Matches local path to mounted sync folder
- Constructs SharePoint URL from UrlNamespace

### Config.bas
Centralized configuration.
- Template path constants (10 template filenames)
- `GetBasePath()` - OneDrive folder via USERPROFILE env var
- `GetTemplatePath()` - Full path to specific template
- `GetExportRoot()` - Output folder for generated workbooks

## VBA/Excel Constraints Handled

### VBA Trim() Limitation
VBA's `Trim()` function only removes ASCII space (char 32). It does NOT remove non-breaking spaces (char 160) which commonly appear in web copy/paste. Custom sanitization handles this.

### FormatCondition.Borders API
When applying conditional formatting, the `Borders` property doesn't support full styling options. Must apply border formatting separately from conditional formatting rules.

### Sheet Event Binding
VBA standard modules cannot directly bind to worksheet events. The `InstallChangeHandler` function injects event handler code into Sheet1's code module programmatically using `VBComponents`.

### Hyperlinks.Add Side Effect
`Hyperlinks.Add` automatically changes the font color to blue and adds underline. Must explicitly restore desired font styling after adding hyperlinks.

## Debug Tips

All modules log to the Immediate Window (Ctrl+G in VBA editor) with prefixed tags:
- `[Config]` - Path resolution
- `[Dispatcher]` - Order type routing
- `[EntryPoint]` - UI setup and transaction flow
- `[FileHelper]` - Template operations
- `[PathHelper]` - Path construction
- `[SharePointHelper]` - URL generation
- `[TX_NewUsage]`, `[TX_Return]`, `[TX_Swap]` - Form building and data mapping
