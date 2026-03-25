Attribute VB_Name = "DisableCloseX"
'Attribute VB_Name = "DisableCloseX"
'Include this code at the top of the module
Private Const GWL_STYLE = -16
Private Const WS_CAPTION = &HC00000
Private Const WS_SYSMENU = &H80000
Private Const SC_CLOSE = &HF060

#If VBA7 Then

    Private Declare PtrSafe Function FindWindowA _
        Lib "user32" (ByVal lpClassName As String, _
        ByVal lpWindowName As String) As Long
    Private Declare PtrSafe Function DeleteMenu _
        Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, _
        ByVal wFlags As Long) As Long
    Private Declare PtrSafe Function GetSystemMenu _
        Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
        
#Else

    Private Declare Function FindWindowA _
        Lib "user32" (ByVal lpClassName As String, _
        ByVal lpWindowName As String) As Long
    Private Declare Function DeleteMenu _
        Lib "user32" (ByVal hMenu As Long, _
        ByVal nPosition As Long, ByVal wFlags As Long) As Long
    Public Declare Function GetSystemMenu _
        Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
        
#End If

'Include this code in the same module as the API calls above
Public Sub CloseButtonSettings(frm As Object, show As Boolean)

Dim windowHandle As Long
Dim menuHandle As Long
windowHandle = FindWindowA(vbNullString, frm.Caption)

If show = True Then

    menuHandle = GetSystemMenu(windowHandle, 1)

Else

    menuHandle = GetSystemMenu(windowHandle, 0)
    DeleteMenu menuHandle, SC_CLOSE, 0&

End If

End Sub

