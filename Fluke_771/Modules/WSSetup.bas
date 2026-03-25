Attribute VB_Name = "WSSetup"
Sub SetupWS()
    ' Disable events temporarily to avoid recursive triggers
    Application.EnableEvents = False
    

    ' Change the Make, Model, and Description Here
    make = "Fluke"
    Model = "789"
    UnitDesc = "Processmeter"

    ' Tab Names
    Tab1 = "Datasheet"
    Tab2 = ""
    Tab3 = ""
    Tab4 = ""
    
    'tabName = Tab1
    
    ' Accredited Tab Names
    ATab1 = "Accredited"
    ATab2 = ""
    ATab3 = ""
    ATab4 = ""

    ' Column Letters (for example, F and G)
    'PreloadCols = Split("F G")  <- Example Usage
    PreloadCols = Split("F G")

    
    'You get this number after you setup your array and then run the
    'Preload Button
    'DR = 113
    'This number is vital in retrieving your data from the database
    
    '''''''''''''''''''' No editing below this point ''''''''''''''''''''''
    ' Workbook Initialization
    Dim sourceWorkbook As Workbook
    Dim ThisWorkbookPath As String
    
    ' === Workbook and Worksheet Setup ===
Dim wbPath As String
Dim wbName As String
Dim wb As Workbook
Dim wsTab1 As Worksheet
Dim wsTab2 As Worksheet
Dim wsTab3 As Worksheet
Dim wsTab4 As Worksheet
Dim wsATab1 As Worksheet
Dim wsATab2 As Worksheet
Dim wsATab3 As Worksheet
Dim wsATab4 As Worksheet

' Get the path of the current workbook
wbPath = ThisWorkbook.Path
wbName = ThisWorkbook.Name
Set wb = ThisWorkbook

 'Set references for Tab(NonAccredited)


'Set references for Tab (NonAccredited)
If Len(Trim(Tab1)) > 0 Then Set wsTab1 = GetSheetIfExists(Tab1)
If Len(Trim(Tab2)) > 0 Then Set wsTab2 = GetSheetIfExists(Tab2)
If Len(Trim(Tab3)) > 0 Then Set wsTab3 = GetSheetIfExists(Tab3)
If Len(Trim(Tab4)) > 0 Then Set wsTab4 = GetSheetIfExists(Tab4)

'Set references for Tab (Accredited)
If Len(Trim(ATab1)) > 0 Then Set wsATab1 = GetSheetIfExists(ATab1)
If Len(Trim(ATab2)) > 0 Then Set wsATab2 = GetSheetIfExists(ATab2)
If Len(Trim(ATab3)) > 0 Then Set wsATab3 = GetSheetIfExists(ATab3)
If Len(Trim(ATab4)) > 0 Then Set wsATab4 = GetSheetIfExists(ATab4)


    ' === End Workbook and Worksheet Setup ===
    
 
 
    ThisWorkbookPath = ThisWorkbook.Path
    Set sourceWorkbook = ThisWorkbook
    Set InfoSheet = sourceWorkbook.Sheets("Information")
    Set InfoDataSheet = sourceWorkbook.Sheets("Datasheet")

    ' Setup WorkOrderSheet
    Set WorkOrderSheet = InfoSheet
    Set dataSheet = InfoDataSheet

    ' Read Calibrator Model and DMM Model
    'CalibratorModel = WorkOrderSheet.Range("M9").Value
    'DMMModel = WorkOrderSheet.Range("O10").Value

    ' WorkOrder Number Cell Location
    Set cellAddress = WorkOrderSheet.Range("H13")

    ' WorkOrder Value
    WorkOrder = WorkOrderSheet.Range("H13").Value

    ' Fill in WorkOrder information
    WorkOrderSheet.Range("X3").Value = make
    WorkOrderSheet.Range("Y3").Value = Model
    WorkOrderSheet.Range("W4").Value = UnitDesc

    ' Re-enable events after the code runs
    Application.EnableEvents = True
End Sub

Function GetSheetIfExists(SheetName As String, Optional wb As Workbook) As Worksheet
    If wb Is Nothing Then Set wb = ThisWorkbook

    If Len(Trim(SheetName)) = 0 Then Exit Function

    On Error Resume Next
    Set GetSheetIfExists = wb.Sheets(SheetName)
    On Error GoTo 0
End Function



