Attribute VB_Name = "DatasheetCode"
Public Sub TestArray()
    Dim i As Long
    Dim startRow As Long
    Dim endRow As Long
    Dim currentRow As Long
    Dim activeCol As Long
    
    activeCol = ActiveCell.Column
    currentRow = ActiveCell.Row

    If Not IsArray(ranges) Or IsEmpty(ranges) Then
        ArraySetup
    End If

    If activeCol >= ColNumAF And activeCol <= ColNumAL Then
        For i = LBound(ranges) To UBound(ranges)
            startRow = CLng(Split(ranges(i), ":")(0))
            endRow = CLng(Split(ranges(i), ":")(1))

            If currentRow >= startRow And currentRow <= endRow Then
                TestSect = i + 1  ' 1-based index
                MsgBox TestSect
                On Error Resume Next
                ActiveCell.offset(1, 0).Select
                If err.Number <> 0 Then
                    MsgBox "Error selecting next cell: " & err.Description
                    err.Clear
                End If
                On Error GoTo 0
                Exit Sub
            End If
        Next i
    End If
End Sub

Sub CheckIfActiveCellInValidRangeDynamic()
    Dim startRow As Long, endRow As Long
    Dim validRangeF As Range, validRangeG As Range, fullValidRange As Range

    ' Make sure ranges array is initialized
    If Not IsArray(ranges) Or IsEmpty(ranges) Then ArraySetup

    ' Get dynamic start and end row from ranges array
    GetMinAndMaxRowsFromArray ranges, startRow, endRow

    ' Define column letters (F and G)
    Dim colLetterF As String: colLetterF = "F"
    Dim colLetterG As String: colLetterG = "G"

    ' Build valid column ranges
    Set validRangeF = Range(colLetterF & startRow & ":" & colLetterF & endRow)
    Set validRangeG = Range(colLetterG & startRow & ":" & colLetterG & endRow)

    ' Union of F and G ranges
    Set fullValidRange = Union(validRangeF, validRangeG)

    ' Check active cell location
    If Not Intersect(ActiveCell, fullValidRange) Is Nothing Then
        'MsgBox "You're inside the valid range! Proceed as normal."
    Else
        CommToggle "Standby"
        PrevTestSect = 0
        PrevSameTest = 0
        OffValueV = 0
        OffValueU = ""
        HVImageShow
        'MsgBox "You're outside the valid range. Cut code off!"
    End If
End Sub


Public Sub HandleSelectionChange(ByVal Target As Excel.Range)

On Error GoTo ErrorHandler
    Dim colRange As Boolean
    Dim isInRange As Boolean
    Dim ColLetter As String
    Dim cellAddress As String
    Dim foundValidRange1 As Boolean
    Dim foundValidRange2 As Boolean
    Dim foundValidRange3 As Boolean
    Dim rangePair As Variant
    Dim btn As Shape
    Dim i As Long
    Dim startRow As Long
    Dim endRow As Long
    Dim currentRow As Long
    Dim btnExists As Boolean
    Dim shp As Shape
    'Dim tag As String
    
    ' Get active cell info
    cellAddress = ActiveCell.Address(0, 0) ' No absolute references
    ColLetter = Left(cellAddress, 1)       ' Gets first letter (assumes columns A–Z)

btnExists = False
For Each shp In Sheet2.Shapes
    If shp.Name = "CommToggle" Then
        btnExists = True
        Set btn = shp
        Exit For
    End If
Next shp

If Not btnExists Then
    MsgBox "CommToggle button not found on Sheet2.", vbExclamation
    Exit Sub
End If

    activeCol = ActiveCell.Column
    
    SetupWS
    currentRow = ActiveCell.Row
    inRange = False

    If Not IsArray(ranges) Or IsEmpty(ranges) Then
        ArraySetup
    End If

If Sheet2.Range("AA1").Value = "Standby" Or Sheet2.Range("AA1").Value = "Operating" Then

            
    ''' Main Test Section for Test Points
    
    If activeCol >= ColNumAF And activeCol <= ColNumAL Then
        For i = LBound(ranges) To UBound(ranges)
            startRow = CLng(Split(ranges(i), ":")(0))
            endRow = CLng(Split(ranges(i), ":")(1))

            If currentRow >= startRow And currentRow <= endRow Then
                With btn
                .TextFrame.Characters.Text = "Operating"
                With .TextFrame.Characters.Font
                    .Size = 20
                    .Bold = True
                    .Color = RGB(255, 0, 0)
                End With
            End With
                Sheet2.Range("AA1").Value = "Operating"
                TestSect = i + 1  ' 1-based index
                GetValues
                HVImageShow
                TestOp
                'MsgBox TestSect
                'MsgBox "3OffValueV: " & OffValueV & " OffValueU: " & OffValueU & " OffValueHz: " & OffValueHz & " OffValueHzU: " & OffValueHzU & " Offset: " & offset & " OffSetU: " & OffSetU & " Wave: " & Wave & " Duty: " & duty

                On Error Resume Next
               ' ActiveCell.offset(1, 0).Select
                If err.Number <> 0 Then
                    MsgBox "Error selecting next cell: " & err.Description
                    err.Clear
                End If
                On Error GoTo 0
                Exit Sub
            End If
        Next i
    'End If
    
    ''' End Test Section for Test Points
    
    ''' Skipping normal blank spaces
    
    'If activeCol >= ColNumAF And activeCol <= ColNumAL Then
    For i = LBound(Skips) To UBound(Skips)
        startRow = CLng(Split(Skips(i), ":")(0))
        endRow = CLng(Split(Skips(i), ":")(1))
        
        If currentRow >= startRow And currentRow <= endRow Then
            ' Move down one cell only
            On Error Resume Next
            'MsgBox PrevSameTest
            Skipped = True
            ActiveCell.offset(1, 0).Select
            If err.Number <> 0 Then
                MsgBox "Error selecting next cell: " & err.Description
                err.Clear
            End If
            On Error GoTo 0
            Exit Sub  ' Exit after moving once
        End If
    Next i
    'Exit Sub
    'End If
    
    ''' End Skipping normal Blank spaces
    
    ''' Skipping Blank spaces and putting calibrator in standby
    
    'If activeCol >= ColNumAF And activeCol <= ColNumAL Then
    For i = LBound(stdbyComms) To UBound(stdbyComms)
        startRow = CLng(Split(stdbyComms(i), ":")(0))
        endRow = CLng(Split(stdbyComms(i), ":")(1))
        
        If currentRow >= startRow And currentRow <= endRow Then
            On Error Resume Next
            'ShowImageInCell "HVImage", "AA1"
            CommToggle "Standby"
            'TestSect = 0
            'PrevSameTest = 0
            Comm False, True, False
            'MsgBox PrevSameTest
            Cls
            Skipped = True
            ActiveCell.offset(1, 0).Select
            If err.Number <> 0 Then
                MsgBox "Error selecting next cell: " & err.Description
                err.Clear
            End If
            On Error GoTo 0
            Exit Sub  ' Exit after moving once
        End If
    Next i
    Exit Sub
    End If
    
    ''' End Skipping Blank spaces and putting calibrator in standby

        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '-------------------End Here on last Gray Cell---------------------
    If ActiveCell.Address = LastCellF Or ActiveCell.Address = LastCellG Or ActiveCell.Address = LastCellH Then
   ' MsgBox PrevSameTest
    CommToggle "Standby"
    Comm False, True, False
    ActiveSheet.Range("I9").Select
    TestSect = 0
    PrevSameTest = 0
    MsgBox PrevSameTest
    End If
    
    '''''''''''''''''''''''''''Clicking anywhere else'''''''''''''''''''''''''''

'CheckIfActiveCellInValidRangeDynamic
'SkipsAndComms
'SkipComms
'If Skipped = True Then
'MsgBox "Skipped"
'Skipped = False
'Exit Sub
'End If
CommToggle "Standby"
    Comm False, True, False
TestSect = 0
PrevSameTest = 0
SameTest = 0
'MsgBox "OutSide Area: SameTest: " & SameTest & " PrevSameTest: " & PrevSameTest & " TestSect: " & TestSect
Else




End If
        

 

    Exit Sub

ErrorHandler:
    Application.EnableEvents = True
    MsgBox "Error: " & err.Description
    Call ReportError("YourMacroName", err.Number, err.Description, Erl)

End Sub
Sub ReportError(procName As String, ErrNum As Long, ErrDesc As String, ErrLine As Long)
    MsgBox "Error in " & procName & vbCrLf & _
           "Line: " & ErrLine & vbCrLf & _
           "Error " & ErrNum & ": " & ErrDesc, vbCritical
End Sub


Function GetMinAndMaxRowsFromArray(arr As Variant, ByRef minRow As Long, ByRef maxRow As Long)
    Dim i As Long
    Dim startRow As Long, endRow As Long
    minRow = 999999
    maxRow = 0

    For i = LBound(arr) To UBound(arr)
        startRow = CLng(Split(arr(i), ":")(0))
        endRow = CLng(Split(arr(i), ":")(1))

        If startRow < minRow Then minRow = startRow
        If endRow > maxRow Then maxRow = endRow
    Next i
End Function

