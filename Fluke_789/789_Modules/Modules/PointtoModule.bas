Attribute VB_Name = "PointtoModule"
'Sub OpenModule(ByVal ModuleName As String)

    ' Ensure VBE is visible
 '   Application.VBE.MainWindow.Visible = True
    
    ' Activate the code window for the given module
  '  Application.VBE.VBProjects(ThisWorkbook.VBProject.Name) _
   '     .VBComponents(ModuleName).CodeModule.CodePane.show

'End Sub
Sub OpenModule(moduleName As String)
    Dim vbComp As VBComponent
    Dim found As Boolean
    
    On Error Resume Next
    
    ' Make sure the VBA editor is visible
    Application.VBE.MainWindow.Visible = True
    
    ' Look only in the active workbook's VBProject
    For Each vbComp In ActiveWorkbook.VBProject.VBComponents
        If vbComp.Name = moduleName Then
            vbComp.CodeModule.CodePane.show
            found = True
            Exit For
        End If
    Next vbComp
    
    If Not found Then
        MsgBox "Module or UserForm '" & moduleName & "' was not found in " & ActiveWorkbook.Name, vbExclamation
    End If
End Sub


'Sub OpenUserFormCode(FormName As String)
 '   Application.VBE.MainWindow.Visible = True
  '  Application.VBE.VBProjects(ThisWorkbook.VBProject.Name) _
   '     .VBComponents(FormName).CodeModule.CodePane.show
'End Sub
Sub OpenUserFormCode(FormName As String)
    Dim vbComp As VBComponent
    Dim found As Boolean
    
    On Error Resume Next
    
    ' Make sure the VBA editor is visible
    Application.VBE.MainWindow.Visible = True
    
    ' Look only in the active workbook's VBProject
    For Each vbComp In ActiveWorkbook.VBProject.VBComponents
        If vbComp.Type = vbext_ct_MSForm Then ' Only check UserForms
            If vbComp.Name = FormName Then
                vbComp.CodeModule.CodePane.show
                found = True
                Exit For
            End If
        End If
    Next vbComp
    
    If Not found Then
        MsgBox "UserForm '" & FormName & "' was not found in " & ActiveWorkbook.Name, vbExclamation
    End If
End Sub



