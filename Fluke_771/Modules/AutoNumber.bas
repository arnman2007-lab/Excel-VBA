Attribute VB_Name = "AutoNumber"
Sub AddLineNumbersToProcedure()
    Dim codeMod As CodeModule
    Dim procName As String
    Dim startLine As Long
    Dim procLines As Long
    Dim i As Long
    Dim currentLine As String
    
    Set codeMod = ThisWorkbook.VBProject.VBComponents("DatasheetCode").CodeModule ' Change to your module name
    procName = "HandleSelectionChange" ' Change to your sub name

    startLine = codeMod.ProcStartLine(procName, vbext_pk_Proc)
    procLines = codeMod.ProcCountLines(procName, vbext_pk_Proc)

    For i = startLine To startLine + procLines - 1
        currentLine = codeMod.Lines(i, 1)
        If Trim(currentLine) <> "" And Left(Trim(currentLine), 1) <> "'" Then
            codeMod.ReplaceLine i, "'" & i & ": " & currentLine
        End If
    Next i

    MsgBox "Line numbers added as comments for debugging."
End Sub

