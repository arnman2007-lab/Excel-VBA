VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PanelForm 
   Caption         =   "Control Panel"
   ClientHeight    =   6585
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3870
   OleObjectBlob   =   "PanelForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "PanelForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CodeButton_Click()
    Dim currentState As String
    
    If ToggleStates Is Nothing Then InitStates
    currentState = ToggleStates("CodeButton")
    
    Select Case currentState
        Case "Off"
            ButtonState Me, "CodeButton", "Standby"
            'PanelForm.STDAction.Caption = "Standby"
            'DoEvents   ' Forces the label to repaint immediately
        Case "Standby"
            ButtonState Me, "CodeButton", "Off"
            'PanelForm.STDAction.Caption = "Off"
            'DoEvents   ' Forces the label to repaint immediately
        Case "Operating"
            ButtonState Me, "CodeButton", "Standby"
            CalibClearStatus "Standby"
            'PanelForm.STDAction.Caption = "Standby"
            'DoEvents   ' Forces the label to repaint immediately
            
    End Select
    
End Sub

Private Sub DSPrint_Click()
    SetupWS
        
If Range("J8").Value = "Status: Incomplete" Or WorkOrderSheet.Range("H14").Value = "" Or WorkOrderSheet.Range("H15").Value = "" Or WorkOrderSheet.Range("H16").Value = "" Then
MsgBox "Please Fix Empty Cells - Check spaces between input cells too"
        
Exit Sub
End If
        
    PrintSelection.show
    
End Sub

Private Sub GetData_Click()
    Dim tabName     As String
    tabName = ActiveSheet.Name
    
    RetrieveData tabName
    
End Sub



Private Sub OpenEditPanel_Click()
        On Error Resume Next
    If EditPanelForm Is Nothing Then
        Set EditPanelForm = New EditPanelForm
    End If
    EditPanelForm.show vbModeless
    
End Sub

Private Sub PreloadStuff_Click()
    Dim tabName     As String
    tabName = Tab1
    
    Preload
    
End Sub

Private Sub ResetDatasheet_Click()
    Dim tabName     As String
    tabName = ActiveSheet.Name
    
    ResetCells tabName
    
End Sub

Private Sub StoreInputData_Click()
    Dim tabName     As String
    tabName = ActiveSheet.Name
    
    StoreData tabName
    
End Sub

Private Sub UserForm_Initialize()
DockPanelOnRight Me
    Dim savedState As String
    Dim EditStatus As Boolean
    ' Ensure dictionary exists
    InitStates
   EditStatus = True
    
    ' Read state from helper cell
    On Error Resume Next
    savedState = ThisWorkbook.Sheets("Information").Range("QQ1").Value
    On Error GoTo 0
    
    If savedState = "" Then savedState = "Off"   ' default if empty
    
    ' Apply saved state to button (updates caption, color, and dictionary)
    ButtonState Me, "CodeButton", savedState
    
    CodeButton.TakeFocusOnClick = False
    DSPrint.TakeFocusOnClick = False
    ResetDatasheet.TakeFocusOnClick = False
    StoreInputData.TakeFocusOnClick = False
    GetData.TakeFocusOnClick = False
    If EditStatus = True Then
    PreloadStuff.TakeFocusOnClick = False
    OpenEditPanel.TakeFocusOnClick = False
    PreloadStuff.Visible = True
    OpenEditPanel.Visible = True
    
    ElseIf EditStatus = False Then
    PreloadStuff.TakeFocusOnClick = False
    OpenEditPanel.TakeFocusOnClick = False
    PreloadStuff.Visible = False
    OpenEditPanel.Visible = False
    End If
    
    
    ' Start manually positioned
    Me.StartUpPosition = 0
    
    ' Initial position: docked to the right if checkbox checked
   ' DockPanel True
    CalibratorModel = WorkOrderSheet.Range("M9").Value
    DMMModel = WorkOrderSheet.Range("P9").Value
    CounterModel = WorkOrderSheet.Range("M16").Value
    make = WorkOrderSheet.Range("X3").Value
    Model = WorkOrderSheet.Range("Y3").Value
    UnitDesc = WorkOrderSheet.Range("W4").Value
    
    Me.ModelLabel.Caption = make
    Me.ModelLabel.Caption = Model
    Me.UnitDescLabel.Caption = UnitDesc
    Me.CalibratorLabel.Caption = CalibratorModel
    Me.DMMLabel.Caption = DMMModel
    Me.CounterLabel.Caption = CounterModel
    
End Sub

' -------------------------------
' Dock / Undock Logic
' -------------------------------
Private Sub chkDock_Click()
    DockPanel chkDock.Value
End Sub

Private Sub DockPanel(DockRight As Boolean)
    With Me
        If DockRight Then
            ' Dock to the right of usable Excel window
            .Left = Application.UsableWidth - .Width - 10
            .Top = 50
        Else
            ' Free float: current position stays
            .Left = .Left
            .Top = .Top
        End If
    End With
End Sub



