# Excel-VBA Calibration Projects

Excel-based calibration automation templates with VBA macros for various test equipment.

## Workstation Setup

Before running any calibration, configure your workstation:

1. Open `Special_WorkBooks/1WorkStation Setup.xlsm`
2. Click "Get device identification" to query all connected GPIB devices
3. Assign each GPIB address to the correct standard (Calibrator, DMM, Counter)
4. This generates `DeviceInfo.csv` with your workstation's specific configuration

Each calibration workbook reads from this central `DeviceInfo.csv`.

## Starting a New Automation Project

### Initial Setup
1. Create new project folder (e.g., `Fluke_87V/`)
2. Copy `Special_WorkBooks/BackEnd Template.xlsm` to the new folder
3. Rename it for your DUT (e.g., `Fluke 87V Multimeter.xlsm`)
4. Create `Modules/`, `UserForms/`, `Images/` subfolders
5. Copy VBA modules from a similar existing project

### Required Module Customizations

| Module | What to Customize |
|--------|-------------------|
| **WSSetup.bas** | `make`, `Model`, `UnitDesc` - MUST match the DUT |
| **SetupArrays.bas** | `ranges(Tab1)` - row ranges for each test section |
| **DatasheetCode.bas** | `Select Case Target.Address` - cell addresses and calibrator commands |
| **TestSectionNumbers.bas** | Test section cases (1000=operational, 6000=mA, etc.) |

### Required Sheet Code (.cls files in Sheets/ folder)

| File | Sheet | Purpose |
|------|-------|---------|
| **ThisWorkbook.cls** | ThisWorkbook | `Workbook_Open` - loads DeviceInfo.csv, initializes standards |
| **Sheet1.cls** | Datasheet | `Worksheet_SelectionChange` → `HandleSelectionChange` |
| **wsInfo.cls** | Information | `Worksheet_Change` (trims), button handlers |
| **Sheet5.cls** | Buttons And Code | Buttons to open modules in VBA editor |

Import these `.cls` files via File → Import in VBA Editor.

### Final Steps
1. Import all `.bas` and `.frm` files in Excel VBA Editor
2. Add the ThisWorkbook and Datasheet sheet code above
3. Add hookup diagrams to `Images/[DUT Name]/[Calibrator Model]/`
4. Test with workstation setup configured

## Projects

| Folder | Description |
|--------|-------------|
| `Special_WorkBooks/` | Shared tools: workstation setup, backend template, DeviceInfo.csv |
| `Examiner_1000/` | Monarch Examiner 1000 Vibration Meter calibration |
| `Fluke_789/` | Fluke 789 Processmeter calibration |
| `Fluke_771/` | Fluke 771 Milliamp Process Clamp Meter calibration |

## Project Structure

Each project folder contains:
- `.xlsm` files - Excel workbooks with VBA macros
- `Modules/` - Exported VBA module files (.bas)
- `UserForms/` - Exported VBA userform files (.frm, .frx)
- `Sheets/` - Exported sheet/workbook code (.cls)
- `Images/` - Equipment hookup diagrams
