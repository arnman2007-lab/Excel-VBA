Attribute VB_Name = "DataSave"


Sub StoreData(Model As String)

    Dim lastRow As Long
    Dim WO As Variant
    Dim foundCell As Range
    Dim j As Long
    Dim dataStoragePath As String
    Dim dataStorageWorkbook As Workbook
    Dim DataStorage As Worksheet
    Dim i As Integer
    Dim rangePair As Variant
    Dim startRow As Long
    Dim endRow As Long
    Dim rangeSplit() As String
    SetupWS
    If Not IsArray(ranges) Or IsEmpty(ranges) Then
    ArraySetup
End If

    ' Define the path to the DataStorage workbook
    dataStoragePath = ThisWorkbook.Path & "\Database\DataStorage.xlsx"

    ' Check if the Database folder exists, if not, create it
    If Dir(ThisWorkbook.Path & "\Database", vbDirectory) = "" Then
        MkDir ThisWorkbook.Path & "\Database"
    End If

    ' Check if the file exists, if not, create it
    If Dir(dataStoragePath) = "" Then
        ' Create the workbook
        Set dataStorageWorkbook = Workbooks.Add
        ' Save it to the specified path
        On Error GoTo ErrorHandler
        dataStorageWorkbook.SaveAs fileName:=dataStoragePath
        On Error GoTo 0
        ' Set the DataStorage worksheet
        Set DataStorage = dataStorageWorkbook.Sheets(1)
        ' Optionally, rename the worksheet to "DataStorage"
        DataStorage.Name = "DataStorage"
        ' Close the workbook to finish creating it
        dataStorageWorkbook.Close
    End If

    ' Open the DataStorage workbook
    Set dataStorageWorkbook = Workbooks.Open(dataStoragePath)
    ' Set the DataStorage worksheet
    Set DataStorage = dataStorageWorkbook.Sheets(1) ' Assuming DataStorage is the first sheet

    WO = WorkOrder
    If WO <> "" Then
        searchWOValue = WorkOrder
    Else
        TypeWorkOrder.show
        SetupWS
        WO = WorkOrder
    End If

    ' Search for the value in column B of DataStorage sheet
    Set foundCell = DataStorage.Columns("B").Find(What:=WO, LookIn:=xlValues, LookAt:=xlWhole)
    
    ' If value already exists, delete the entire row
    If Not foundCell Is Nothing Then
        foundCell.EntireRow.Delete
    End If

    If Model = Tab1 Then
        lastRow = Application.WorksheetFunction.Min(DataStorage.Cells(DataStorage.Rows.Count, "B").End(xlUp).Row + 1, 10000)
        i = 2

        'Information Tab
        'For j = 13 To 16
          '  DataStorage.Cells(lastRow, i).Value = WorkOrderSheet.Range("H" & j).Value
         '   i = i + 1
        'Next j
            DataStorage.Cells(lastRow, i).Value = WorkOrderSheet.Range("H13").Value
            i = i + 1
            DataStorage.Cells(lastRow, i).Value = WorkOrderSheet.Range("X3").Value
            i = i + 1
            DataStorage.Cells(lastRow, i).Value = WorkOrderSheet.Range("Y3").Value
            i = i + 1
            'DataStorage.Cells(lastRow, i).Value = WorkOrderSheet.Range("W4").Value
            'i = i + 1
            DataStorage.Cells(lastRow, i).Value = WorkOrderSheet.Range("H14").Value
            i = i + 1
            DataStorage.Cells(lastRow, i).Value = WorkOrderSheet.Range("H15").Value
            i = i + 1
            DataStorage.Cells(lastRow, i).Value = WorkOrderSheet.Range("H16").Value
            i = i + 1
        
        
        

        ' Loop through each range and call FillData
    
    For Each rangePair In ranges
        
        rangeSplit = Split(rangePair, ":")
        startRow = CLng(rangeSplit(0))
        endRow = CLng(rangeSplit(1))
    For j = startRow To endRow
    DataStorage.Cells(lastRow, i).Value = dataSheet.Range("F" & j).Value
    i = i + 1
    DataStorage.Cells(lastRow, i).Value = dataSheet.Range("G" & j).Value
    i = i + 1
    Next j
    Next rangePair
    
    ElseIf Model = Tab2 Then
        

    ElseIf Model = Tab3 Then
        

    ElseIf Model = Tab4 Then
        

    End If

    ' Save and close the DataStorage workbook
    dataStorageWorkbook.Save
    dataStorageWorkbook.Close

   ' MsgBox "Data stored successfully."

    Exit Sub

ErrorHandler:
    MsgBox "Error creating or accessing DataStorage file: " & err.Description

End Sub

Sub RetrieveData(tabName As String)
Dim searchValue As String
    Dim foundCell As Range
    Dim lastRow As Long
    Dim copyRange As Range
    Dim destinationRanges() As Range
    Dim i As Long
    Dim j As Long
    Dim dbFolderPath As String
    Dim dataStorageFileName As String
    Dim dataStorageFilePath As String
    Dim externalWB As Workbook
    Dim externalWS As Worksheet
    Dim rangePair As Variant
    Dim startRow As Long
    Dim endRow As Long
    Dim rangeSplit() As String
    
    SetupWS
    If Not IsArray(ranges) Or IsEmpty(ranges) Then
    ArraySetup
End If

    searchValue = WorkOrder
    If searchValue <> "" Then
    searchValue = WorkOrder
    Else
    TypeWorkOrder.show
    SetupWS
    searchValue = WorkOrder
    End If
    
    ' Determine current workbook's folder path
    dbFolderPath = ThisWorkbook.Path
    
    ' Specify the filename of DataStorage.xlsx
    dataStorageFileName = "DataStorage.xlsx"
    
    ' Construct full path to DataStorage.xlsx in the Database subfolder
    dataStorageFilePath = dbFolderPath & "\Database\" & dataStorageFileName
    
    ' Check if DataStorage.xlsx exists
    If Not Dir(dataStorageFilePath) <> "" Then
        MsgBox "DataStorage.xlsx not found in the specified path: " & dbFolderPath & "\Database\"
        Exit Sub
    End If
    
    ' Open the external workbook
    On Error Resume Next
    Set externalWB = Workbooks.Open(dataStorageFilePath)
    On Error GoTo 0
    
    If externalWB Is Nothing Then
        MsgBox "Error opening DataStorage.xlsx. Please check the file path and try again."
        Exit Sub
    End If
    
    ' Set the worksheet in the external workbook
    Set externalWS = externalWB.Worksheets("DataStorage")
    If externalWS Is Nothing Then
        externalWB.Close False
        Set externalWB = Nothing
        MsgBox "Worksheet 'DataStorage' not found in DataStorage.xlsx. Exiting..."
        Exit Sub
    End If
    
    
   ' Find the cell containing the search value in column B of data sheet
    Set foundCell = externalWS.Columns("B:B").Find(What:=searchValue, LookIn:=xlValues, LookAt:=xlWhole)
    
    ' Check if the value is found
    If Not foundCell Is Nothing Then
        ' If found, get the row number
        lastRow = foundCell.Row
       
        ' Define the range to copy (entire row)
        Set copyRange = externalWS.Rows(lastRow)
        
        
    If tabName = Tab1 Then
        ReDim destinationRanges(1 To DR)
        i = 1
        
    'Information Tab
    For j = 14 To 16
        Set destinationRanges(i) = WorkOrderSheet.Range("H" & j)
        
        i = i + 1
    Next j
       ' Set destinationRanges(i) = WorkOrderSheet.Range("H13")
        '    i = i + 1
         '   Set destinationRanges(i) = WorkOrderSheet.Range("X3")
          '  i = i + 1
           ' Set destinationRanges(i) = WorkOrderSheet.Range("Y3")
            'i = i + 1
       '     Set destinationRanges(i) = WorkOrderSheet.Range("W4")
        '    i = i + 1
         '   Set destinationRanges(i) = WorkOrderSheet.Range("H14")
          '  i = i + 1
           ' Set destinationRanges(i) = WorkOrderSheet.Range("H15")
            'i = i + 1
      '      Set destinationRanges(i) = WorkOrderSheet.Range("H16")
       '     i = i + 1
    
    For Each rangePair In ranges
        
        rangeSplit = Split(rangePair, ":")
        startRow = CLng(rangeSplit(0))
        endRow = CLng(rangeSplit(1))
    For j = startRow To endRow
    Set destinationRanges(i) = dataSheet.Range("F" & j)
    i = i + 1
    Set destinationRanges(i) = dataSheet.Range("G" & j)
    i = i + 1
    Next j
    Next rangePair
    
    
'MsgBox i - 1
    
    
    
    
        ElseIf tabName = Tab2 Then
        
    
    ReDim destinationRanges(1 To 350)
    j = 1
    
   'Information Tab
   For i = 14 To 16
        Set destinationRanges(j) = Worksheets("Information").Range("H" & i)
        
        j = j + 1
    Next i
    
    For Each rangePair In ranges2
        
        rangeSplit = Split(rangePair, ":")
        startRow = CLng(rangeSplit(0))
        endRow = CLng(rangeSplit(1))
    For j = startRow To endRow
    Set destinationRanges(i) = dataSheet.Range("F" & j)
    i = i + 1
    Set destinationRanges(i) = dataSheet.Range("G" & j)
    i = i + 1
    Next j
    Next rangePair
    
        Else
            MsgBox "Invalid tabName specified."
            externalWB.Close False
            Set externalWB = Nothing
            Exit Sub
    
    
    
    End If
    
        ' Loop through the array and paste values
        For i = LBound(destinationRanges) To UBound(destinationRanges)
            If Not destinationRanges(i) Is Nothing Then
            
                'destinationRanges(i).Value = copyRange.Cells(1, i + 2).Value ' Starting from column B (2nd column)
                destinationRanges(i).Value = copyRange.Cells(1, i + 4).Value
                'MsgBox WorkOrderSheet.Range("H13").Value
            Else
            
                MsgBox "Destination range not set for value " & j
            End If
            'j = j + 1 ' Increment j for the next iteration
        Next i
     ' Close the external workbook
        externalWB.Close False
        Set externalWB = Nothing
        
    Else
        MsgBox "WorkOrder not found in DataStorage.xlsx!"
    End If
End Sub


