Attribute VB_Name = "Dictionary"
' --- Module1 ---

Public ToggleStates As Scripting.Dictionary

Public Sub InitStates()
    If ToggleStates Is Nothing Then
        Set ToggleStates = New Scripting.Dictionary
    End If
    
    ' Only add if missing, don't overwrite
    If Not ToggleStates.Exists("CodeButton") Then ToggleStates.Add "CodeButton", "Off"
    'If Not ToggleStates.Exists("AnotherButton") Then ToggleStates.Add "AnotherButton", "Off"
    'If Not ToggleStates.Exists("ThirdButton") Then ToggleStates.Add "ThirdButton", "Off"
End Sub

Public Sub ButtonState(FormObj As Object, btnName As String, state As String)
    Dim btn As MSForms.CommandButton
    On Error GoTo NotFound
    
    Set btn = FormObj.Controls(btnName)
    
    Select Case LCase(state)
        Case "operating"
            btn.Caption = "Operating"
            btn.BackColor = vbRed
                        PanelForm.STDAction.Caption = "Operating"
            DoEvents   ' Forces the label to repaint immediately
            
        Case "standby"
            btn.Caption = "Standby"
            btn.BackColor = vbGreen
                        PanelForm.STDAction.Caption = "Standby"
            DoEvents   ' Forces the label to repaint immediately
            
        Case "off"
            btn.Caption = "Off"
            btn.BackColor = &H8000000F
                        PanelForm.STDAction.Caption = "Off"
            DoEvents   ' Forces the label to repaint immediately
            
        Case Else
            btn.Caption = "Unknown"
            btn.BackColor = vbGray
    End Select
    
    ' Save state in dictionary
    If ToggleStates Is Nothing Then InitStates
    If Not ToggleStates.Exists(btnName) Then ToggleStates.Add btnName, state
    ToggleStates(btnName) = state
    
    ' Save to helper cell on Information tab
    ThisWorkbook.Sheets("Information").Range("QQ1").Value = state
    
    Exit Sub

NotFound:
    MsgBox "Button '" & btnName & "' not found on form.", vbExclamation
End Sub


