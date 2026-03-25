Attribute VB_Name = "SkipswithComms"
Sub SkipsAndComms()
    Dim i As Long
    Dim startRow As Long
    Dim endRow As Long
    Dim currentRow As Long
If Not IsArray(Skips) Or IsEmpty(Skips) Then
    ArraySetup
End If

    ' Get current row
    currentRow = ActiveCell.Row
    
    ' Use regular For loop instead of For Each
If activeCol >= ColNumAF And activeCol <= ColNumAL Then
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
End If
End Sub

Sub SkipComms()
    Dim i As Long
    Dim startRow As Long
    Dim endRow As Long
    Dim currentRow As Long
    
If Not IsArray(stdbyComms) Or IsEmpty(stdbyComms) Then
    ArraySetup
End If
    ' Get current row
    currentRow = ActiveCell.Row
    
    ' Use regular For loop instead of For Each
If activeCol >= ColNumAF And activeCol <= ColNumAL Then
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
End If
End Sub


