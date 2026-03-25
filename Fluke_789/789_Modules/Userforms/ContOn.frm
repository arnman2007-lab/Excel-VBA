VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ContOn 
   Caption         =   "Continuity On"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "ContOn.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ContOn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Fail_Click()


ActiveCell.Value = "Fail"
Unload Me
Selection.offset(1, 0).Select

End Sub

Private Sub Terminate_Click()
ActiveCell.offset(0, -2).Select
Unload Me
PrevAddress = ActiveCell.Address
shouldContinue = False

End Sub

Private Sub Pass_Click()

ActiveCell.Value = "Pass"
Unload Me
Selection.offset(1, 0).Select

End Sub

Private Sub UserForm_Initialize()
 CenterUserFormOnActiveSheet Me
Call CloseButtonSettings(Me, False)

End Sub
