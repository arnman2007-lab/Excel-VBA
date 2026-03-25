Attribute VB_Name = "Communication"

Sub Comm(shouldInitialize As Boolean, shouldStandby As Boolean, shouldControlRen As Boolean)



    If shouldInitialize Then
    
        If InitializedCalibrator = "True" Then
    
        Else
            InitializedCalibrator = "False"
           ' MsgBox InitializedCalibrator
        If Worksheets("Information").Range("Calibrator") = "" Then
    
        Else
            CalibListed = True
            Set addstdvalue = Worksheets("Information").Range("Calibrator")
            Set ioMgr = New VisaComLib.ResourceManager
            Set instrument = New VisaComLib.FormattedIO488
            Set instrument.IO = ioMgr.Open(addstdvalue.Value)
            InitializedCalibrator = "True"
        End If
    End If
    
        If Initialized3458 = "True" Then
    
        Else
            Initialized3458 = "False"
           ' MsgBox Initialized3458
        If Worksheets("Information").Range("DMM") = "" Then
    
        Else
            Set add3458value = Worksheets("Information").Range("DMM")
            Set ioMgr = New VisaComLib.ResourceManager
            Set inst3458 = New VisaComLib.FormattedIO488
            Set inst3458.IO = ioMgr.Open(add3458value.Value)
            Initialized3458 = "True"
        End If
    End If
    
    Else
    End If
    
    
    
    If shouldStandby Then
    
        If Worksheets("Information").Range("Calibrator") = "" Then
    
        Else
            instrument.WriteString "*RST"
           'instrument.WriteString "*cls"
           ' instrument.WriteString "STBY"
            
        End If
    
    Else
    End If
    
    
    
    If shouldControlRen Then
    
        If Worksheets("Information").Range("DMM") = "" Then
    
        Else
            'If func = "dci" Or func = "dcv" Then
            Set gpib3458 = ioMgr.Open(inst3458.IO.resourceName)
            inst3458.WriteString "RESET"
            gpib3458.ControlRen GPIB_REN_GTL
            gpib3458.Close
        
            ' Else
             ' End If
        
       End If
       
       If Worksheets("Information").Range("Calibrator") = "" Then
    
        Else
            Set gpib = ioMgr.Open(instrument.IO.resourceName)
            gpib.ControlRen GPIB_REN_GTL
            gpib.Close
        End If
    
    Else
    End If
    

End Sub


Sub Cls()



    
    
    
    
   If Worksheets("Information").Range("Calibrator") = "" Then
    
        Else
           ' instrument.WriteString "*RST"
           'instrument.WriteString "*cls"
            instrument.WriteString "STBY"
            instrument.WriteString "*cls"
            instrument.WriteString "OUT " & 0 & " " & "mV" & ", " & 0 & " " & "hz" ' & "; OPER"
        End If
    
    
   
    
    
  

End Sub



Sub ListGPIBAddresses1()
    On Error Resume Next

    Dim resources() As String
    Dim resourceCount As Long
    Dim RM_P As VisaComLib.ResourceManager
    Dim resourceName As Variant ' Declare resourceName as Variant
    Dim gpibResources() As String
    Dim asrlResources() As String
    Dim infoSheet As Worksheet
    Dim calibratorCell As Range
    Dim dmmCell As Range
    Dim commCell As Range
    Dim i As Long

    ' Set reference to the Information sheet
    Set infoSheet = ThisWorkbook.Sheets("Information")
    ' Set the cells where you want to populate the dropdown lists
    Set calibratorCell = infoSheet.Range("Calibrator")
    Set dmmCell = infoSheet.Range("DMM")
    Set commCell = infoSheet.Range("Comm")

    ' Clear existing content in the cells
    calibratorCell.ClearContents
    dmmCell.ClearContents
    commCell.ClearContents

    Set RM_P = New VisaComLib.ResourceManager

    ' Get all available resources
    resources = RM_P.FindRsrc("?*")

    ' Loop through resources and filter for GPIB devices and ASRL devices, excluding "GPIB0::INTFC"
    For Each resourceName In resources ' No need to change this line
        If InStr(1, resourceName, "GPIB") > 0 And InStr(1, resourceName, "INTFC") = 0 Then
            resourceCount = resourceCount + 1
            ReDim Preserve gpibResources(1 To resourceCount)
            gpibResources(resourceCount) = resourceName
        ElseIf InStr(1, resourceName, "ASRL") > 0 Then
            resourceCount = resourceCount + 1
            ReDim Preserve asrlResources(1 To resourceCount)
            asrlResources(resourceCount) = resourceName
        End If
    Next resourceName

    Set RM_P = Nothing

    ' Display the list of GPIB devices in dropdown lists
    If Not IsEmpty(gpibResources) Then
        ' Populate dropdown list in the Calibrator cell
        With calibratorCell.Validation
            .Delete ' Clear existing validation
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
                 xlBetween, Formula1:=Join(gpibResources, ",")
            .IgnoreBlank = True
            .InCellDropdown = True
            .ShowInput = True
            .ShowError = True
        End With
        
        ' Populate dropdown list in the DMM cell
        With dmmCell.Validation
            .Delete ' Clear existing validation
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
                 xlBetween, Formula1:=Join(gpibResources, ",")
            .IgnoreBlank = True
            .InCellDropdown = True
            .ShowInput = True
            .ShowError = True
        End With
    Else
        MsgBox "No GPIB devices found."
    End If

    ' Display the list of ASRL devices in dropdown list for the Comm cell
    If Not IsEmpty(asrlResources) Then
        ' Populate dropdown list in the Comm cell
        With commCell.Validation
            .Delete ' Clear existing validation
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
                 xlBetween, Formula1:=Join(asrlResources, ",")
            .IgnoreBlank = True
            .InCellDropdown = True
            .ShowInput = True
            .ShowError = True
        End With
    Else
        MsgBox "No ASRL devices found."
    End If
End Sub

Sub ListGPIBAddresses()
    On Error Resume Next

    Dim resources() As String
    Dim RM_P As VisaComLib.ResourceManager
    Dim resourceName As Variant ' Declare resourceName as Variant
    Dim gpibResources As String
    Dim asrlResources As String
    Dim infoSheet As Worksheet
    Dim calibratorCell As Range
    Dim dmmCell As Range
    Dim commCell As Range

    ' Set reference to the Information sheet
    Set infoSheet = ThisWorkbook.Sheets("Information")
    ' Set the cells where you want to populate the dropdown lists
    Set calibratorCell = infoSheet.Range("Calibrator")
    Set dmmCell = infoSheet.Range("DMM")
    Set commCell = infoSheet.Range("Comm")

    ' Clear existing content in the cells
    calibratorCell.ClearContents
    dmmCell.ClearContents
    commCell.ClearContents

    Set RM_P = New VisaComLib.ResourceManager

    ' Get all available resources
    resources = RM_P.FindRsrc("?*")

    ' Loop through resources and filter for GPIB devices and ASRL devices, excluding "GPIB0::INTFC"
    For Each resourceName In resources
        If InStr(1, resourceName, "GPIB") > 0 And InStr(1, resourceName, "INTFC") = 0 Then
            gpibResources = gpibResources & "," & resourceName
        ElseIf InStr(1, resourceName, "ASRL") > 0 Then
            asrlResources = asrlResources & "," & resourceName
        End If
    Next resourceName

    Set RM_P = Nothing

    ' Display the list of GPIB devices in dropdown lists
    If Len(gpibResources) > 0 Then
        ' Populate dropdown list in the Calibrator cell
        With calibratorCell.Validation
            .Delete ' Clear existing validation
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
                 xlBetween, Formula1:="," & Mid(gpibResources, 2)
            .IgnoreBlank = True
            .InCellDropdown = True
            .ShowInput = True
            .ShowError = True
        End With
        
        ' Populate dropdown list in the DMM cell excluding the selected value from the Calibrator cell
        With dmmCell.Validation
            .Delete ' Clear existing validation
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
                 xlBetween, Formula1:="," & Mid(gpibResources, 2)
            .IgnoreBlank = True
            .InCellDropdown = True
            .ShowInput = True
            .ShowError = True
        End With
        
        ' Populate dropdown list in the Comm cell
        With commCell.Validation
            .Delete ' Clear existing validation
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
                 xlBetween, Formula1:=asrlResources
            .IgnoreBlank = True
            .InCellDropdown = True
            .ShowInput = True
            .ShowError = True
        End With
    Else
        MsgBox "No GPIB devices found."
    End If
End Sub


