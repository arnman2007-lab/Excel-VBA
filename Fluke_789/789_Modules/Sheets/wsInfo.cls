VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "wsInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub GetCommPorts_Click()
ListCommPorts
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
   
    
    Dim cell As Range
    Application.EnableEvents = False
    For Each cell In Target
        If Not IsEmpty(cell.Value) And Not cell.HasFormula Then
            cell.Value = Trim(cell.Value)
        End If
    Next cell
    Application.EnableEvents = True
    
End Sub




