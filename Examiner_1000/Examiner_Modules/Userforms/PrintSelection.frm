VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PrintSelection 
   Caption         =   "Print Selection"
   ClientHeight    =   4965
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7335
   OleObjectBlob   =   "PrintSelection.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "PrintSelection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





Private Sub PrintDatasheet_Click()
    Dim currentSheetName As String
    currentSheetName = ActiveSheet.Name ' Get the name of the active sheet
    
    'Print the; Current; Sheet
    PrintTab currentSheetName
    
    
    Unload Me
End Sub




Private Sub PrintAccredited_Click()
Dim currentSheetName As String
currentSheetName = ActiveSheet.Name

    PrintTabAcred currentSheetName
    Unload Me
End Sub

Private Sub UserForm_Initialize()
    CenterUserFormOnActiveSheet Me
End Sub

