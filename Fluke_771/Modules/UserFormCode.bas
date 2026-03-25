Attribute VB_Name = "UserFormCode"
Sub UForms(UForm As String)
    ' Make sure no leftover instance is sitting in memory
    If IsFormLoaded("MainHookup") Then
        Unload MainHookup
    End If
    
    Select Case UForm
        Case "MainForm", "MainForm_Sourcing", "MainForm_Basic_Comment"
            MainHookup.show
    End Select
End Sub

Function IsFormLoaded(frmName As String) As Boolean
    Dim frm As Object
    For Each frm In VBA.UserForms
        If frm.Name = frmName Then
            IsFormLoaded = True
            Exit Function
        End If
    Next frm
    IsFormLoaded = False
End Function

