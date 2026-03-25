Attribute VB_Name = "DataSave"
Sub StoreData(tabName As String)
    Dim lastRow As Long
    Dim WO As Variant
    Dim foundCell As Range
    Dim j As Long, i As Long, c As Long
    Dim dataStoragePath As String
    Dim dataStorageWorkbook As Workbook
    Dim DataStorage As Worksheet
    Dim rangePair As Variant
    Dim startRow As Long, endRow As Long
    Dim rangeSplit() As String
    Dim arr As Variant
    Dim colLetter As String
    Dim wsSource As Worksheet
    
    ' Setup worksheets, ranges, etc.
    SetupWS
    ArraySetup
    
    ' Get the source sheet
    On Error GoTo SheetNotFound
    Set wsSource = ThisWorkbook.Worksheets(tabName)
    On Error GoTo 0
    
    ' Get the correct ranges array for the sheet
    Select Case tabName
        Case Tab1
            If ranges.Exists(Tab1) Then arr = ranges(Tab1)
        Case Tab2
            If ranges.Exists(Tab2) Then arr = ranges(Tab2)
        Case Tab3
            If ranges.Exists(Tab3) Then arr = ranges(Tab3)
        Case Tab4
            If ranges.Exists(Tab4) Then arr = ranges(Tab4)
        Case Else
            MsgBox "No ranges defined for this sheet: " & tabName, vbExclamation
            Exit Sub
    End Select
    
    If Not IsArrayInitialized(arr) Then
        MsgBox "No valid ranges for sheet: " & tabName, vbExclamation
        Exit Sub
    End If
    
    ' Create/open storage workbook
    dataStoragePath = ThisWorkbook.Path & "\Database\DataStorage.xlsx"
    If Dir(ThisWorkbook.Path & "\Database", vbDirectory) = "" Then MkDir ThisWorkbook.Path & "\Database"
    If Dir(dataStoragePath) = "" Then
        Set dataStorageWorkbook = Workbooks.Add
        dataStorageWorkbook.SaveAs Filename:=dataStoragePath
        Set DataStorage = dataStorageWorkbook.Sheets(1)
        DataStorage.Name = "DataStorage"
        dataStorageWorkbook.Close
    End If
    Set dataStorageWorkbook = Workbooks.Open(dataStoragePath)
    Set DataStorage = dataStorageWorkbook.Sheets(1)
    
    ' Handle WorkOrder
    WO = WorkOrder
    If WO = "" Then
        TypeWorkOrder.show
        SetupWS
        WO = WorkOrder
    End If
    
    ' Delete existing WorkOrder row
    Set foundCell = DataStorage.Columns("B").Find(What:=WO, LookIn:=xlValues, LookAt:=xlWhole)
    If Not foundCell Is Nothing Then foundCell.EntireRow.Delete
    
    ' Determine last row to write
    lastRow = Application.WorksheetFunction.Min(DataStorage.Cells(DataStorage.Rows.Count, "B").End(xlUp).Row + 1, 10000)
    i = 2 ' Start column
    
    ' Optional: fixed info for Tab1
    If tabName = Tab1 Then
        DataStorage.Cells(lastRow, i).Value = WorkOrderSheet.Range("H13").Value: i = i + 1
        DataStorage.Cells(lastRow, i).Value = WorkOrderSheet.Range("X3").Value: i = i + 1
        DataStorage.Cells(lastRow, i).Value = WorkOrderSheet.Range("Y3").Value: i = i + 1
        DataStorage.Cells(lastRow, i).Value = WorkOrderSheet.Range("H14").Value: i = i + 1
        DataStorage.Cells(lastRow, i).Value = WorkOrderSheet.Range("H15").Value: i = i + 1
        DataStorage.Cells(lastRow, i).Value = WorkOrderSheet.Range("H16").Value: i = i + 1
    End If
    
    ' Loop through the sheet-specific ranges
    Dim k As Long, cellValue As Variant
    For k = LBound(arr) To UBound(arr)
        rangeSplit = Split(arr(k), ":")
        startRow = CLng(rangeSplit(0))
        endRow = CLng(rangeSplit(1))
        
        For j = startRow To endRow
            For c = LBound(PreloadCols) To UBound(PreloadCols)
                colLetter = PreloadCols(c)
                cellValue = wsSource.Range(colLetter & j).Value
                DataStorage.Cells(lastRow, i).Value = cellValue
                i = i + 1
            Next c
        Next j
    Next k
    
    ' Save and close
    dataStorageWorkbook.Save
    dataStorageWorkbook.Close
    Exit Sub

SheetNotFound:
    MsgBox "Sheet '" & tabName & "' not found."
End Sub

Sub RetrieveData(tabName As String)
    Dim searchValue As String
    Dim foundCell As Range
    Dim dataStoragePath As String
    Dim dataStorageWorkbook As Workbook
    Dim DataStorage As Worksheet
    Dim lastRow As Long
    Dim i As Long, j As Long, c As Long
    Dim arr As Variant
    Dim rangeSplit() As String
    Dim startRow As Long, endRow As Long
    Dim k As Long
    Dim colLetter As String
    Dim wsSource As Worksheet

    SetupWS
    ArraySetup
    
    searchValue = WorkOrder
    If searchValue = "" Then
        TypeWorkOrder.show
        SetupWS
        searchValue = WorkOrder
    End If

    ' Open DataStorage
    dataStoragePath = ThisWorkbook.Path & "\Database\DataStorage.xlsx"
    If Dir(dataStoragePath) = "" Then
        MsgBox "DataStorage.xlsx not found at " & ThisWorkbook.Path & "\Database\"
        Exit Sub
    End If
    
    Set dataStorageWorkbook = Workbooks.Open(dataStoragePath)
    Set DataStorage = dataStorageWorkbook.Sheets("DataStorage")
    
    ' Find WorkOrder row
    Set foundCell = DataStorage.Columns("B").Find(What:=searchValue, LookIn:=xlValues, LookAt:=xlWhole)
    If foundCell Is Nothing Then
        MsgBox "WorkOrder '" & searchValue & "' not found in DataStorage."
        dataStorageWorkbook.Close False
        Exit Sub
    End If
    lastRow = foundCell.Row
    
    ' Get sheet-specific array
    Set wsSource = ThisWorkbook.Worksheets(tabName)
    Select Case tabName
        Case Tab1
            If ranges.Exists(Tab1) Then arr = ranges(Tab1)
        Case Tab2
            If ranges.Exists(Tab2) Then arr = ranges(Tab2)
        Case Tab3
            If ranges.Exists(Tab3) Then arr = ranges(Tab3)
        Case Tab4
            If ranges.Exists(Tab4) Then arr = ranges(Tab4)
        Case Else
            MsgBox "No ranges defined for sheet: " & tabName
            dataStorageWorkbook.Close False
            Exit Sub
    End Select
    
    If Not IsArrayInitialized(arr) Then
        MsgBox "No valid ranges for sheet: " & tabName
        dataStorageWorkbook.Close False
        Exit Sub
    End If
    
    i = 2 ' start reading from column 2
    
    ' Optional fixed info for Tab1
    If tabName = Tab1 Then
        WorkOrderSheet.Range("H13").Value = DataStorage.Cells(lastRow, i): i = i + 1
        WorkOrderSheet.Range("X3").Value = DataStorage.Cells(lastRow, i): i = i + 1
        WorkOrderSheet.Range("Y3").Value = DataStorage.Cells(lastRow, i): i = i + 1
        WorkOrderSheet.Range("H14").Value = DataStorage.Cells(lastRow, i): i = i + 1
        WorkOrderSheet.Range("H15").Value = DataStorage.Cells(lastRow, i): i = i + 1
        WorkOrderSheet.Range("H16").Value = DataStorage.Cells(lastRow, i): i = i + 1
    End If
    
    ' Loop through the sheet-specific ranges
    For k = LBound(arr) To UBound(arr)
        rangeSplit = Split(arr(k), ":")
        startRow = CLng(rangeSplit(0))
        endRow = CLng(rangeSplit(1))
        
        For j = startRow To endRow
            For c = LBound(PreloadCols) To UBound(PreloadCols)
                colLetter = PreloadCols(c)
                wsSource.Range(colLetter & j).Value = DataStorage.Cells(lastRow, i).Value
                i = i + 1
            Next c
        Next j
    Next k
    
    dataStorageWorkbook.Close False
End Sub

