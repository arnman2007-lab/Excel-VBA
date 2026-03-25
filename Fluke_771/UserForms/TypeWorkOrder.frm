VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} TypeWorkOrder 
   Caption         =   "Input WorkOrder Number"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "TypeWorkOrder.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "TypeWorkOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False








Private Sub SendWONum_Click()
    ' Ensure that the Information sheet exists
    On Error GoTo ErrorHandler
    'Dim ws As Worksheet
    'Set ws = WorkOrderSheet
    
    ' Transfer the value from TextBox1 to H13 on the Information sheet
    cellAddress = Me.TextBox1.Value
    Unload Me
    Exit Sub

ErrorHandler:
    MsgBox "The 'Information' sheet does not exist.", vbExclamation
End Sub



Private Sub UserForm_Initialize()
    CenterUserFormOnActiveSheet Me
End Sub
