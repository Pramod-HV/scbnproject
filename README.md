# MXL File Processor

Automated tool for processing Sterling B2Bi MXL map files with multiple features including ExplicitRule insertion, character encoding modification, PreSession rules, and more.

## Quick Start

### For ZIP files (extracts and renames):
```bash
python mxl_parser.py zipfiles
```

### For direct MXL files (no renaming):
```bash
python mxl_parser.py old_mxlFiles --mxl

### Process All MXL Files (Recommended)
```bash
python process_all_mxl_files.py mapping_results.xlsx Generic_checklistMain.xlsm
```

### Process Single MXL File
```bash
python process_data_updater_v3.py mapping_results.xlsx Generic_checklistMain.xlsm mxl_files/YOUR_FILE.mxl
```

### Fix Corrupted MXL Files
```bash
# Remove namespace prefixes (ns0:, ns1:, etc.)
python remove_namespace_prefixes.py mxl_files/
```

### Extract Codelists Names
```bash
python3 codelist_extractor.py
```

### Rename Codelists Names
```bash
python3 codelist_renamer.py
```

### Extract Excel
```bash
python process_data_manager.py extract mxl_files process_data_rules.xlsx
```

### Comment/Uncomment
```bash
python process_data_manager.py apply process_data_rules.xlsx mxl_files
```

---

## Project Overview

This project automates the processing of Sterling B2Bi EDI MXL (map) files for migration and standardization. It reads configuration from Excel files and applies various transformations to MXL files based on business rules.

## Features

### 1. ExplicitRule Insertion
- Automatically inserts `<ExplicitRule>` tags with SQL update statements
- Matches map name from MXL file with `mapping_results.xlsx`
- Applies rules from `Generic_checklistMain.xlsm` (process_data_updates sheet)
- Supports EDI (X12, EDIFACT) and NON-EDI formats (XML, Positional, CSV)
- Handles field positions like BHT-03, N104(N1*ST), XML paths

### 2. Character Encoding Modification
- Modifies INPUT/OUTPUT `<CharacterEncoding>` tags
- Reads configuration from Generic_rule sheet
- Options: "Keep as is", "NONE" (empty tag), or specific encoding value
- Special handling: Skips INPUT encoding for XML input maps

### 3. PreSession Rule Updates
- Updates PreSession rules with Base Name from Generic_rule sheet
- Triggered when "Pre-session rule" = "Yes"
- Integrates with existing PreSession rule logic

### 4. X Syntax Token Modification
- Adds extended character range to X syntax token
- Triggered when "Modify the X syntax token" = "Yes"
- Ensures X12 compliance

### 5. Freeformat Processing
- Processes freeformat sections (BothSides, InputOnly, OutputOnly)
- Triggered based on "Freeformat" setting in Generic_rule sheet
- Converts fixed-format to freeformat

### 6. Comment/Uncomment Management
- **Phase 1 (Extract)**: Extracts comment status from MXL files to Excel
- **Phase 2 (Apply)**: Applies comment status from Excel back to MXL files
- Allows bulk comment/uncomment operations via Excel editing
- Preserves user comments during processing

### 7. Inbound Maps Positional Delimiter Feature
- Configures positional delimiter settings for inbound maps
- Reads configuration from Inbound_maps sheet
- Updates `<PosSyntax>` section with delimiter values (CR, LF, etc.)
- Handles DelimiterUsed and Delimiter2Used flags

### 8. Error Reporting
- Generates Excel error reports (error_report.xlsx)
- Tracks inactive fields with no active segments
- Identifies output fields with no input mapping
- Provides detailed error information for troubleshooting

### 9. Inactive Field Handling
- Smart handling of inactive EDI segments
- Searches for next active field in same segment
- Falls back gracefully when no active fields found

## Configuration Files

### mapping_results.xlsx
Contains map metadata with columns:
- **Map Name**: Name extracted from MXL file Description tag
- **Direction**: Inbound or Outbound
- **Document Type**: 850, 856, DELINS, ORDERS, etc.
- **Input Format**: Format of input (e.g., "EDI (Edifact, DELINS, 003)")
- **Output Format**: Format of output (e.g., "EDI(X12, 850, 004010)")

### Generic_checklistMain.xlsm
Contains processing rules with multiple sheets:
- **Generic_rule**: Global settings (encoding, PreSession, freeformat, X syntax token)
- **process_data_updates**: ExplicitRule insertion rules (Direction, Document Type, Field Position, XPath)
- **Inbound_maps**: Positional delimiter configuration

## Field Position Formats

The tool supports various field position formats:

- **EDI Segment-Field**: `BHT-03` (Segment "BHT", Field 3)
- **EDI with Qualifier**: `N104(N1*ST)` (Segment "N1", Field 4, Qualifier "ST")
- **XML Path**: `SPI_ASN/Header/ShipmentID` (Hierarchical XML path)
- **Document Type Prefix**: `850.BSN-02` (Document type 850, Segment BSN, Field 2)

## Document Type Extraction

The tool extracts document type from Input/Output Format columns:

**Priority Order:**
1. Input Format (3 elements): `EDI (Edifact, DELINS, 003)` → "DELINS"
2. Input Format (2 elements): `EDI(X12, 850)` → "850"
3. Input Format (1 element): `Positional (SALESORDER)` → "SALESORDER"
4. Output Format (3 elements): `EDI(X12, 850, 004010)` → "850"
5. Output Format (2 elements): `EDI(X12, 850)` → "850"
6. Output Format (1 element): `XML (DOCTYPE)` → "DOCTYPE"
7. Fallback: 3-digit pattern search

**Special Case:** "SAP" is skipped as it's a marker, not a document type.

## Requirements

```bash
pip install -r requirements.txt
```

Required packages:
- pandas
- openpyxl

## Project Structure

```
process_data_updates/
├── process_all_mxl_files.py          # Batch processor (recommended)
├── process_data_updater_v3.py        # Single file processor
├── modify_character_encoding.py      # Encoding modifier
├── update_presession_rules.py        # PreSession rule updater
├── modify_syntax_token.py            # X syntax token modifier
├── mxl_processor.py                  # Freeformat processor
├── inbound_mapsFeatures.py           # Inbound maps positional delimiter
├── remove_namespace_prefixes.py      # Fix namespace corruption
├── error_reporter.py                 # Error reporting utility
├── inactive_field_reporter.py        # Inactive field reporter
├── handle_inactive_fields.py         # Inactive field handler
├── process_data_manager.py           # Comment/uncomment manager
├── comment_uncomment.py              # Comment extraction/application
├── mapping_results.xlsx              # Map metadata
├── Generic_checklistMain.xlsm        # Processing rules
└── mxl_files/                        # MXL files directory
```

## Workflow

1. **Preparation**: Place MXL files in `mxl_files/` directory
2. **Configuration**: Update `mapping_results.xlsx` and `Generic_checklistMain.xlsm`
3. **Processing**: Run `process_all_mxl_files.py` to process all files
4. **Review**: Check `error_report.xlsx` for any issues
5. **Validation**: Review modified MXL files

## Output

- Modified MXL files saved in place
- Log files created for each run
- Error reports generated (error_report.xlsx)
- Terminal output shows progress and results

## Notes

- Always backup your MXL files before processing
- Review Generic_checklistMain.xlsm settings before running
- Use dry-run mode to preview changes: `--dry-run`
- Character encoding: Self-closing tags are functionally equivalent to opening/closing tags
- Use remove_namespace_prefixes.py if files get corrupted with ns0: prefixes
- The tool preserves user comments during processing

## Troubleshooting

### Files corrupted with namespace prefixes (ns0:, ns1:)
```bash
python remove_namespace_prefixes.py mxl_files/
```

### Error reports
Check `error_report.xlsx` for:
- Inactive fields with no active segments
- Output fields with no mapping
- Fields requiring manual review

### Document type not matching
- Verify Input Format/Output Format in mapping_results.xlsx
- Check if document type exists in process_data_updates sheet
- Review log output for document type extraction details
