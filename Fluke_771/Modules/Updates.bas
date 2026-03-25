Attribute VB_Name = "Updates"
Sub CheckAndUpdateModules()
    Dim networkPath As String
    Dim localPath As String
    Dim fso As Object
    Dim srcFile As Object
    Dim destFile As Object
    Dim file As Object
    Dim srcFolder As Object
    Dim destFolder As Object
    Dim fileName As String
    
    ' Set network path
    networkPath = "I:\Location - Alexandria\Alexandria Lab\EXCEL_SHEETS\Paul Excel\Addins\289\Modules"
    
    ' Set local path (same folder as the workbook in "Addins" subfolder)
    localPath = ThisWorkbook.Path & "\Addins"
    
    ' Create FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Ensure local folder exists
    If Not fso.FolderExists(localPath) Then fso.CreateFolder localPath
    
    ' Check if network path exists
    If fso.FolderExists(networkPath) Then
        ' Network is available ? Use network version
        Set srcFolder = fso.GetFolder(networkPath)
        Set destFolder = fso.GetFolder(localPath)
        
        ' Loop through all .bas files in the network folder
        For Each file In srcFolder.Files
            If LCase(fso.GetExtensionName(file.Name)) = "bas" Then
                fileName = file.Name
                
                ' Check if the file exists in the local folder
                If fso.FileExists(localPath & "\" & fileName) Then
                    ' Compare timestamps
                    Set srcFile = fso.GetFile(networkPath & "\" & fileName)
                    Set destFile = fso.GetFile(localPath & "\" & fileName)
                    
                    If srcFile.DateLastModified > destFile.DateLastModified Then
                        ' Newer file found ? Copy and Replace
                        fso.CopyFile srcFile.Path, destFile.Path, True
                      '  MsgBox "Updated module: " & fileName, vbInformation, "Update Successful"
                    End If
                Else
                    ' File does not exist ? Copy it
                    fso.CopyFile file.Path, localPath & "\" & fileName, True
                    'MsgBox "New module added: " & fileName, vbInformation, "Module Added"
                End If
            End If
        Next
    Else
        ' Network not available, using local files
        MsgBox "Network path not found. Using local modules.", vbExclamation, "Offline Mode"
    End If
    
    ' Cleanup
    Set fso = Nothing
    Set srcFolder = Nothing
    Set destFolder = Nothing
End Sub

Sub RunExternalCode(moduleName As String, subName As String, ParamArray args() As Variant)
    Dim modulePath As String
    Dim vbProj As Object
    Dim comp As Object
    Dim moduleFound As Boolean
    Dim ws As Worksheet
    Dim lastRow As Long
    
    ' Define module path
    modulePath = ThisWorkbook.Path & "\Addins\" & moduleName & ".bas"

    ' Ensure the module file exists
    If Dir(modulePath) = "" Then
        MsgBox "Module not found: " & modulePath, vbExclamation
        Exit Sub
    End If

    ' Get VBA project reference
    Set vbProj = ThisWorkbook.VBProject

    ' Check if module already exists
    moduleFound = False
    For Each comp In vbProj.VBComponents
        If comp.Name = moduleName Then
            moduleFound = True
            Exit For
        End If
    Next comp

    ' If module is not found, import it and log its name
    If Not moduleFound Then
        vbProj.VBComponents.Import modulePath
        
        ' Log module name in HelperSheet
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets("HelperSheet")
        If ws Is Nothing Then
            ' Create HelperSheet if it doesn't exist
            Set ws = ThisWorkbook.Sheets.Add
            ws.Name = "HelperSheet"
            ws.Visible = xlSheetVeryHidden
        End If
        On Error GoTo 0

        ' Find last empty row and store the module name
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
        ws.Cells(lastRow, 1).Value = moduleName
    End If

    ' Run the specified subroutine
    Select Case UBound(args)
        Case -1: Application.Run moduleName & "." & subName
        Case 0: Application.Run moduleName & "." & subName, args(0)
        Case 1: Application.Run moduleName & "." & subName, args(0), args(1)
        Case 2: Application.Run moduleName & "." & subName, args(0), args(1), args(2)
        Case Else: MsgBox "Too many arguments passed.", vbExclamation
    End Select
End Sub


Sub RemoveModule(moduleName As String)
    Dim vbProj As Object
    Dim comp As Object

    ' Get VBA project reference
    Set vbProj = ThisWorkbook.VBProject

    ' Loop through all components and delete the specified module
    On Error Resume Next
    vbProj.VBComponents.Remove vbProj.VBComponents(moduleName)
    On Error GoTo 0
End Sub


