Attribute VB_Name = "ButtonCode"




'----------------The Buttons on the Datasheet point to these subs---------------------


Public Sub DSPrint_Click()
    SetupWS
        
If Range("J8").Value = "Status: Incomplete" Or WorkOrderSheet.Range("H14").Value = "" Or WorkOrderSheet.Range("H15").Value = "" Or WorkOrderSheet.Range("H16").Value = "" Then
MsgBox "Please Fix Empty Cells - Check spaces between input cells too"
        
Exit Sub
End If
        
    PrintSelection.show

End Sub

Public Sub ResetDatasheet_Click()
    Dim tabName     As String
    tabName = ActiveSheet.Name
    
    ResetCells tabName
    
End Sub

Public Sub GetData_Click()
    Dim tabName     As String
    tabName = ActiveSheet.Name
    
    RetrieveData tabName
End Sub

Public Sub SetINOP_Click()
    Dim tabName As String
    tabName = ActiveSheet.Name
    
    Inop tabName
End Sub

Public Sub StoreInputData_Click()
    Dim tabName     As String
    tabName = ActiveSheet.Name
    
    StoreData tabName
    
End Sub

Public Sub PreloadStuff_Click()
    Dim tabName     As String
    tabName = Tab1
    
    Preload tabName
    
End Sub

Option Explicit

Public Sub ShowPanel()
   ' ' Open modeless so sheet stays active
   ' 'PanelForm.show vbModeless
    '    On Error Resume Next
    'If PanelForm Is Nothing Then
     '   Set PanelForm = New PanelForm
    'End If
    'PanelForm.show vbModeless
    ''DoEvents
    ''PositionPanel PanelForm
        On Error Resume Next   ' optional to avoid “already loaded” errors
    
    ' Load the form into memory (so its controls exist)
    Load PanelForm
    
    ' Show modeless so sheet stays active
    If Not PanelForm.Visible Then
        PanelForm.show vbModeless
    End If
    
    ' Optional: position on right side of Excel
    ' PositionPanel PanelForm
End Sub

Sub ClosePanel()
    On Error Resume Next
    Unload PanelForm
End Sub

Public Sub EditShowPanel()
    ' Open modeless so sheet stays active
    'PanelForm.show vbModeless
        On Error Resume Next
    If EditPanelForm Is Nothing Then
        Set EditPanelForm = New EditPanelForm
    End If
    EditPanelForm.show vbModeless
End Sub

Sub CloseEditPanel()
    On Error Resume Next
    Unload EditPanelForm
End Sub

Sub OpenControlPanel()
ShowPanel

End Sub


Public Sub ShowButtonState(btnName As String)
    If ToggleStates Is Nothing Then
        MsgBox "States not initialized yet.", vbExclamation
        Exit Sub
    End If
    
    If ToggleStates.Exists(btnName) Then
        MsgBox btnName & " is currently: " & ToggleStates(btnName), vbInformation
    Else
        MsgBox "No state stored for " & btnName, vbExclamation
    End If
End Sub

