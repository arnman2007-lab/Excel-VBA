VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_Open()

  Call SetupWS
 '-------------------Begin Button Assignment------------------
InitStates
 '-------------------End Button Assignment--------------------

        
DMMReset = 0
    

'Click dropdown...
wsInfo.Range("P15").Value = "Click dropdown..."

    Call ArraySetup
    

   
    HVImageShow 0, "V"
    
   
    

    ' Clear fields
    On Error Resume Next
    wsInfo.Range("M8").Value = ""
    wsInfo.Range("M9").Value = ""
    wsInfo.Range("M10").Value = ""
    wsInfo.Range("M11").Value = ""
    wsInfo.Range("M12").Value = ""
    wsInfo.Range("P8").Value = ""
    wsInfo.Range("P9").Value = ""
    wsInfo.Range("P10").Value = ""
    wsInfo.Range("P11").Value = ""
    wsInfo.Range("M15").Value = ""
    wsInfo.Range("M16").Value = ""
    wsInfo.Range("M17").Value = ""
    wsInfo.Range("M18").Value = ""
    wsInfo.Range("P15").Value = ""
    wsInfo.Range("P16").Value = ""
    wsInfo.Range("P17").Value = ""
    wsInfo.Range("P18").Value = ""
    wsInfo.Range("P19").Value = ""
    wsInfo.Range("P20").Value = ""
    If err.Number <> 0 Then
        MsgBox "Error clearing cell values. Possible merged cell: " & err.Description
        Exit Sub
    End If
    On Error GoTo 0
    


    
    ' Load device info
   
    Call LoadDeviceInfoFromCSV
    

    ' Load into variables
    CalibratorMake = wsInfo.Range("M8").Value
    CalibratorModel = wsInfo.Range("M9").Value
    CalibratorSN = wsInfo.Range("M10").Value
    CalibratorGPIB = wsInfo.Range("M11").Value
    CalibratorScopeOption = wsInfo.Range("M12").Value
    DMMMake = wsInfo.Range("P8").Value
    DMMModel = wsInfo.Range("P9").Value
    DMMSN = wsInfo.Range("P10").Value
    DMMGPIB = wsInfo.Range("P11").Value
    CounterMake = wsInfo.Range("M15").Value
    CounterModel = wsInfo.Range("M16").Value
    CounterSN = wsInfo.Range("M17").Value
    CounterGPIB = wsInfo.Range("M18").Value
    
    Call VerifyConnectedGPIBDevices
    
    Call LoadSpecificModules
    

   
    Exit Sub

ErrorHandler:
    MsgBox "Error opening or reading from 1WorkStation Setup.xlsm:" & vbCrLf & err.Description, vbCritical
End Sub


Private Sub Workbook_SheetActivate(ByVal Sh As Object)
    Dim sheetsToClose As Variant
    Dim i As Long
    
    ' List of sheet names where the PanelForm should close
    sheetsToClose = Array("Information", "ChangeLog", "Buttons And Code", ATab1, ATab2, ATab3, ATab4)  ' change these to your sheet names
    
    For i = LBound(sheetsToClose) To UBound(sheetsToClose)
        If Sh.Name = sheetsToClose(i) Then
            On Error Resume Next
            Unload PanelForm   ' closes it completely
            Exit Sub
        Else
        ShowPanel
        End If
    Next i
End Sub

Sub LoadSpecificModules()
    Dim modulesToLoad As Variant
    Dim i As Integer
    Dim moduleFolder As String
    
    ' List the modules you want to load
    modulesToLoad = Array("CalibratorDrivers", "CalibratorSpecifications", "CalibratorStatusClear", "CounterDrivers", "CounterSpecifications", "CounterStatusClear", "DMMDrivers", "DMMSpecifications", "DMMStatusClear")
    
    moduleFolder = ThisWorkbook.Path & "\modules\"
    
    For i = LBound(modulesToLoad) To UBound(modulesToLoad)
        Call ImportModule(CStr(modulesToLoad(i)), moduleFolder)
    Next i
End Sub

Sub ImportModule(moduleName As String, folderPath As String)
    Dim vbProj As Object
    Dim vbComp As Object
    Dim modulePath As String
    
    Set vbProj = ThisWorkbook.VBProject
    modulePath = folderPath & moduleName & ".bas"
    
    ' Check if file exists
    If Dir(modulePath) = "" Then
        Debug.Print "Module file not found: " & modulePath
        Exit Sub
    End If
    
    ' Remove existing module if present
    For Each vbComp In vbProj.VBComponents
        If vbComp.Name = moduleName Then
            vbProj.VBComponents.Remove vbComp
            Exit For
        End If
    Next vbComp
    
    ' Import the module
    vbProj.VBComponents.Import modulePath
    Debug.Print "Loaded module: " & moduleName
End Sub
