VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Sourcing 
   Caption         =   "Source to DMM"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "Sourcing.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Sourcing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Fail_Click()
    
    Unload Me
   ' ActiveCell.offset(1, 0).Select
    
End Sub

Private Sub Terminate_Click()
TerminateClicked = True
Unload Me
    ActiveCell.offset(0, -2).Select
    
End Sub

Private Sub Pass_Click()
    
    Unload Me
    'ActiveCell.offset(1, 0).Select
    
End Sub

Private Sub UserForm_Initialize()
    CenterUserFormOnActiveSheet Me
    Call CloseButtonSettings(Me, False)
GetValues
    Me.Label1.Caption = "Source " & OffValueV & OffValueU
End Sub

