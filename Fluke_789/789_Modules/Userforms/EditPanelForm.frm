VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} EditPanelForm 
   Caption         =   "Edit Code Panel"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "EditPanelForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "EditPanelForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ArrayEdit_Click()
OpenModule "SetupArrays"
End Sub

Private Sub DatasheetCodeEdit_Click()
OpenModule "DatasheetCode"
End Sub

Private Sub MainHookupFormEdit_Click()
OpenUserFormCode "MainHookup"
End Sub


Private Sub WorksheetEdit_Click()
OpenModule "WSSetup"
End Sub

Private Sub UserForm_Initialize()
   'I nitStates
    'CodeButton.Caption = ToggleStates("CodeButton")
   ' AnotherButton.Caption = ToggleStates("AnotherButton")
    'ThirdButton.Caption = ToggleStates("ThirdButton")
    ' Start manually positioned
    Me.StartUpPosition = 0
    
    ' Initial position: docked to the right if checkbox checked
    DockPanel True
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
