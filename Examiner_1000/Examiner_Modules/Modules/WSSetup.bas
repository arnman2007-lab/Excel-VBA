Attribute VB_Name = "WSSetup"
Sub SetupWS()

    'Change the Make Model and Description Here
    Make = "Monarch Instruments"
    Model = "Examiner 1000"
    UnitDesc = "Vibration Analyzer"
    
    ' Tab Names Change Tab1 to the Datasheet's Tab name Datasheet Datasheet-C etc
    Tab1 = "Datasheet-C"
    Tab2 = ""
    Tab3 = ""
    Tab4 = ""
    
    ' Accredited Tab Names Change ATab1 to the Accredited's Tab name Accredited Accredited-C etc
    ATab1 = "Accredited"
    ATab2 = ""
    ATab3 = ""
    ATab4 = ""
    
    'Since every template is different in column usage; ColNumAF and ColNumAL is used for the current datasheet you are working on.
    'Find what columns the As Found and As Left are i.e. F and G   or G and H, or whatever they may be.
    'Count the columns to the As Found  if it is F then 6, the 6th column from A for example. Put 6 in ColNumAF and do the same for the As Left Column.
    'for the below numbers the As Found and Left columns is F and G
    ColNumAF = 7
    ColNumAL = 8
   
    
    
    'DR: This is for Destination Ranges, used in the DataSave module. After you setup the Arrays in the SetupArrays module, pressing Preload will give you this number. Preload button
    'also makes sure the ranges Array was input correctly but putting numbers in the test point input cells. Great way to check resolution as well.
    
    DR = 73
    
    
''''''''''''''''''''No editing below''''''''''''''''''''''''''''''''''
    'Do not do anything with these
    Dim sourceWorkbook As Workbook
    Dim ThisWorkbookPath As String
    
    'Do not do anything with these
    ThisWorkbookPath = ThisWorkbook.Path
    Set sourceWorkbook = ThisWorkbook
    Set infoSheet = sourceWorkbook.Sheets("Information")
    Set InfoDataSheet = sourceWorkbook.Sheets("Datasheet-C")
    
    'Setup WorkOrderSheet 'Do not do anything with these
    Set WorkOrderSheet = infoSheet
    Set dataSheet = InfoDataSheet
    
    'Do not do anything with these
    CalibModel = WorkOrderSheet.Range("M10").Value
    DMMModel = WorkOrderSheet.Range("M14").Value
    
    ' WorkOrder Number Cell Location 'Do not do anything with these
    Set cellAddress = WorkOrderSheet.Range("H13")
    
    ' WorkOrder Value 'Do not do anything with these
    WorkOrder = WorkOrderSheet.Range("H13").Value
    
    'Do not do anything with these
    WorkOrderSheet.Range("X3").Value = Make
    WorkOrderSheet.Range("Y3").Value = Model
    WorkOrderSheet.Range("W4").Value = UnitDesc
    
    ColLetterAF = Split(Cells(1, ColNumAF).Address(True, False), "$")(0)
    ColLetterAL = Split(Cells(1, ColNumAL).Address(True, False), "$")(0)
    'MsgBox ColLetterAF
    'MsgBox ColLetterAL
    
   
    
End Sub
