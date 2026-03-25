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

1. Copy `Special_WorkBooks/BackEnd Template.xlsm` to a new project folder
2. Rename it for your DUT (e.g., `Fluke 87V Multimeter.xlsm`)
3. Copy VBA modules from an existing project as a starting point
4. Customize `SetupArrays.bas` for your datasheet row ranges
5. Add hookup diagrams to `Images/`

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
- `Images/` - Equipment hookup diagrams
