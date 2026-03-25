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

## VBA Import Instructions
1. Open the xlsx file in Excel
2. Save As -> Excel Macro-Enabled Workbook (.xlsm)
3. Press Alt+F11 to open VBA Editor
4. File -> Import File -> select all .bas files from Modules/
5. File -> Import File -> select all .frm files from UserForms/
6. Save the workbook
