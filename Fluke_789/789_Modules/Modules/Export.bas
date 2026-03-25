Attribute VB_Name = "Export"
Sub ExportAllUserFormsAndModules()
    Dim vbComp As VBIDE.VBComponent
    Dim exportPath As String
    Dim userformPath As String
    Dim modulePath As String
        
    ' Set the export path for UserForms and Modules
    exportPath = ThisWorkbook.Path & "\VBA_ExportsNew\"
    userformPath = exportPath & "Userforms\"
    modulePath = exportPath & "Modules\"
        
    ' Create directories if they do not exist
    If Dir(exportPath, vbDirectory) = "" Then
        MkDir exportPath
    End If
    If Dir(userformPath, vbDirectory) = "" Then
        MkDir userformPath
    End If
    If Dir(modulePath, vbDirectory) = "" Then
        MkDir modulePath
    End If
        
    ' Loop through all VB components in the project
    For Each vbComp In ThisWorkbook.VBProject.VBComponents
        Select Case vbComp.Type
            Case vbext_ct_MSForm
                ' Export UserForm
                vbComp.Export userformPath & vbComp.Name & ".frm"
            Case vbext_ct_StdModule, vbext_ct_ClassModule
                ' Export Standard Module or Class Module
                vbComp.Export modulePath & vbComp.Name & ".bas"
        End Select
    Next vbComp
        
    ' Confirmation message
    MsgBox "All UserForms and Modules have been exported successfully!", vbInformation
End Sub
    
    

