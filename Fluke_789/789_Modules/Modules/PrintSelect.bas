Attribute VB_Name = "PrintSelect"
Function GetFullPrinter(sPtrName As String) As String
    Dim vPort As Variant
    Dim sSetting As String

    Const REG_KEY As String = "HKEY_CURRENT_USER\Software\Microsoft\Windows NT\CurrentVersion\Devices\"

    sSetting = CreateObject("WScript.Shell").RegRead(REG_KEY & sPtrName)
    vPort = Split(sSetting, ",")
    GetFullPrinter = sPtrName & " on " & vPort(1)
End Function

Function FolderExists(ByVal folderPath As String) As Boolean
    ' Check if a folder exists
    On Error Resume Next
    FolderExists = (GetAttr(folderPath) And vbDirectory) = vbDirectory
    On Error GoTo 0
End Function



Public Sub PrintTab(tabName As String)
    Dim StrPath As String
    Dim sourceWorkbook As Workbook
    Dim PrintWorksheet As Worksheet
    Dim SourceValue As Variant
    Dim ThisWorkbookPath As String
    Dim StrPathTemplate As String
    Dim sP As String
    Dim WO As String
    SetupWS
    'wo = Worksheets("Fluke 9142 Field Metrology Well").Range("N4").Value
   
    sP = "Microsoft Print to PDF"
    
    If tabName = Tab1 Then
        StoreData tabName
        RemoveRowsWithX tabName
    
    ElseIf tabName = Tab2 Then
        StoreData tabName
        RemoveRowsWithX tabName
    
    End If
    
    
    Application.ActivePrinter = GetFullPrinter(sP)

    ThisWorkbookPath = ThisWorkbook.Path
    Set sourceWorkbook = ThisWorkbook
    
    
    Set PrintWorksheet = sourceWorkbook.Sheets(tabName)
    
    ' Use the specified tabName directly
    Set SourceWorksheet = WorkOrderSheet

    ' Use Range N4 value from the specified sheet
    SourceValue = cellAddress

    ' Check if the "PDFs" folder exists, and create it if not
    StrPath = ThisWorkbookPath & "\PDFs\"
    If Not FolderExists(StrPath) Then
        MkDir StrPath
    End If

    If Not IsEmpty(SourceValue) Then
        StrPath = ThisWorkbookPath & "\PDFs\"
        

        ' Export PDF from the specified worksheet
        PrintWorksheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:=StrPath & SourceValue & " " & "Datasheet" & ".pdf", Quality:=xlQualityStandard, _
                                            IncludeDocProperties:=False, IgnorePrintAreas:=False, OpenAfterPublish:=False

        
        Application.DisplayAlerts = False
        
        'ThisWorkbook.Close
        
    Else
        MsgBox "Value in cell N4 is empty or invalid.", vbExclamation
    End If
End Sub



Public Sub PrintTabAcred(tabName As String)
    Dim StrPath As String
    Dim sourceWorkbook As Workbook
    Dim PrintWorksheet As Worksheet
    Dim SourceValue As Variant
    Dim ThisWorkbookPath As String
    Dim StrPathTemplate As String
    Dim sP As String
    Dim Model As String
    Dim WO As String
    
    
    sP = "Microsoft Print to PDF"
    
    ' Assigning model based on tabName
    If tabName = Tab1 Then
        StoreData tabName
        RemoveRowsWithXAcred Accredited
        Model = tabName
    ElseIf tabName = Tab2 Then
        StoreData tabName
        RemoveRowsWithXAcred Accredited
        Model = tabName
    
    Else
        ' Add more conditions as needed
        Model = "" ' Default model if not matched
    End If
    
    
    
    Application.ActivePrinter = GetFullPrinter(sP)

    ThisWorkbookPath = ThisWorkbook.Path
    Set sourceWorkbook = ThisWorkbook
    
    
    ' Creating the name of Accredited tab dynamically
    Set PrintWorksheet = sourceWorkbook.Sheets("Accredited")
    
    ' Check if the "PDFs" folder exists, and create it if not
    StrPath = ThisWorkbookPath & "\PDFs\"
    If Not FolderExists(StrPath) Then
        MkDir StrPath
    End If

    
    
    SourceValue = WorkOrder
    

    If Not IsEmpty(SourceValue) Then
        ' Export PDF from the specified worksheet
        PrintWorksheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:=StrPath & SourceValue & " " & "Accredited" & ".pdf", Quality:=xlQualityStandard, _
                                            IncludeDocProperties:=False, IgnorePrintAreas:=False, OpenAfterPublish:=False

        Application.DisplayAlerts = False
        'ThisWorkbook.Close
    Else
        MsgBox "Value in cell H13 is empty or invalid.", vbExclamation
    End If
End Sub





Sub RemoveRowsWithX(tabName As String)
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long

    ' Set the worksheet variable to the sheet with the given name
    Set ws = ThisWorkbook.Sheets(tabName)

    ' Determine the last row in column A
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    ' Loop through each cell in column A from bottom to top
    For i = lastRow To 1 Step -1
        ' Check if the cell value is "X" or "x" (case-insensitive)
        If UCase(ws.Cells(i, "A").Value) = "X" Then
            ' If "X" is found, delete the entire row
            ws.Rows(i).EntireRow.Delete
        End If
    Next i
End Sub

Sub RemoveRowsWithXAcred(tabNameAcred As String)
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
tabName = ATab1
    ' Set the worksheet variable to the sheet with the given name
    Set ws = ThisWorkbook.Sheets(tabName)

    ' Determine the last row in column A
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    ' Loop through each cell in column A from bottom to top
    For i = lastRow To 1 Step -1
        ' Check if the cell value is "X" or "x" (case-insensitive)
        If UCase(ws.Cells(i, "A").Value) = "X" Then
            ' If "X" is found, delete the entire row
            ws.Rows(i).EntireRow.Delete
        End If
    Next i
End Sub



