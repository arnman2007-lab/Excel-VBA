# Fluke 771 Milliamp Process Clamp Meter

## Overview
Calibration automation for the Fluke 771 mA clamp meter. This is a DC current measurement device that clamps around a conductor to measure milliamps without breaking the circuit.

## Test Structure

### Datasheet (Sheet1)
| Rows | Test Section | Description |
|------|--------------|-------------|
| 14-17 | 1000 | Operational checks (Backlight, Display, Keypad, Spotlight) |
| 20-25 | 6000 | DC Current 20.99 mA range (±4, ±12, ±20 mA) |
| 27-28 | 6000 | DC Current 99.9 mA range (±100 mA) |

### Cell Mapping
| Column | Purpose |
|--------|---------|
| B | Function/Range or Test Description |
| C | Nominal Value (calibrator output) |
| D | Units (mA) |
| F | AS FOUND reading |
| G | AS LEFT reading |
| H | LOW acceptance limit |
| J | HIGH acceptance limit |

### Accredited (Sheet2)
Same test points with uncertainty calculations and T.U.R. (Test Uncertainty Ratio).

## Calibrator Specifications Used
- (3.3 to 33) mA: 0.018% + 0.25 µA
- (33 to 330) mA: 0.018% + 2.5 µA

## Hookup
The calibrator sources DC mA through a wire loop. The Fluke 771 clamps around the loop to measure current.

### Images Required
Place hookup diagrams in `Images/Fluke 771/[calibrator model]/`:
- `mA Main Hookup 5500A.jpg`
- `mA Main Hookup 5502A.jpg`
- `mA Main Hookup 5520A.jpg`

## VBA Setup Instructions

### Step 1: Convert to Macro-Enabled
1. Open the xlsx file in Excel
2. Save As -> Excel Macro-Enabled Workbook (.xlsm)

### Step 2: Import Modules (Alt+F11)
1. File -> Import File -> select all `.bas` files from `Modules/`
2. File -> Import File -> select all `.frm` files from `UserForms/`
3. File -> Import File -> select all `.cls` files from `Sheets/`

### Step 3: Sheet Code Mapping
The `.cls` files in `Sheets/` contain code for specific sheets:

| File | Sheet Name | What It Does |
|------|------------|--------------|
| **ThisWorkbook.cls** | ThisWorkbook | `Workbook_Open` - loads DeviceInfo.csv, initializes everything |
| **Sheet1.cls** | Datasheet | `Worksheet_SelectionChange` → calls `HandleSelectionChange` |
| **wsInfo.cls** | Information | `Worksheet_Change` (trims values), `GetCommPorts_Click` |
| **Sheet5.cls** | Buttons And Code | Button handlers for editing modules |

### Step 4: Save the workbook

## Key Modules to Customize for New Projects

| Module | What to Change |
|--------|----------------|
| `WSSetup.bas` | `make`, `Model`, `UnitDesc` - MUST match the DUT |
| `SetupArrays.bas` | `ranges(Tab1)` - row ranges for test points |
| `DatasheetCode.bas` | `Select Case Target.Address` - cell addresses and calibrator commands |
| `TestSectionNumbers.bas` | Test section cases (1000, 6000, etc.) |
