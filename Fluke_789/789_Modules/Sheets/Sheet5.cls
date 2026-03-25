VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub ArraySetupEdit_Click()
OpenModule "SetupArrays"
End Sub

Private Sub DatasheetEdit_Click()
OpenModule "DatasheetCode"
End Sub

Private Sub ToggleButton1_Click()

End Sub

Private Sub UFMainHookUP_Click()
OpenUserFormCode "MainHookup"
End Sub

Private Sub WSSetupEdit_Click()
OpenModule "WSSetup"
End Sub
