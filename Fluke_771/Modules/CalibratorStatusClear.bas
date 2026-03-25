Attribute VB_Name = "CalibratorStatusClear"

Sub CalibClearStatus(StatusClear As String)
     
    CalibratorGPIB = wsInfo.Range("$M$11").Value
    CalibratorModel = wsInfo.Range("$M$9").Value
'Exit sub if no gpib address

'-----------This removes the High Voltage Image if Reset or Standby is Called.
If StatusClear = "Reset" Or StatusClear = "Standby" Then
    'HVImageShow 0, "V"
    
End If

If CalibratorGPIB = "" Then Exit Sub

    'Set ioMgr = New VisaComLib.ResourceManager
    On Error Resume Next
    Set ioMgr = CreateObject("VisaComLib.ResourceManager")
    On Error GoTo 0
    
    If ioMgr Is Nothing Then
        MsgBox "Unable to create VISA Resource Manager. Please check NI-VISA installation.", vbCritical
    End If
    Set CalibDevice = New VisaComLib.FormattedIO488
    Set CalibDevice.IO = ioMgr.Open(CalibratorGPIB)

Select Case StatusClear

    Case "Close"
    
        Select Case CalibratorModel

            Case "5500A"
                PanelForm.STDAction.Caption = CalibratorModel & " Send Command: Close"
                DoEvents
                Set gpib = ioMgr.Open(CalibDevice.IO.resourceName)
                gpib.ControlREN GPIB_REN_GTL
                gpib.Close
        
            Case "5502A"
                PanelForm.STDAction.Caption = CalibratorModel & " Send Command: Close"
                DoEvents
                Set gpib = ioMgr.Open(CalibDevice.IO.resourceName)
                gpib.ControlREN GPIB_REN_GTL
                gpib.Close
        
            Case "5520A"
                PanelForm.STDAction.Caption = CalibratorModel & " Send Command: Close"
                DoEvents
                Set gpib = ioMgr.Open(CalibDevice.IO.resourceName)
                gpib.ControlREN GPIB_REN_GTL
                gpib.Close
                
        
            Case "5522A"
                PanelForm.STDAction.Caption = CalibratorModel & " Send Command: Close"
                DoEvents
                Set gpib = ioMgr.Open(CalibDevice.IO.resourceName)
                gpib.ControlREN GPIB_REN_GTL
                gpib.Close
        
            Case "M3001"
                PanelForm.STDAction.Caption = CalibratorModel & " Send Command: Close"
                DoEvents
                Set gpib = ioMgr.Open(CalibDevice.IO.resourceName)
                gpib.ControlREN GPIB_REN_GTL
                gpib.Close
        
            End Select

    Case "Clear"
    
        Select Case CalibratorModel

            Case "5500A"
                PanelForm.STDAction.Caption = CalibratorModel & " Send Command: Clear"
                DoEvents
                CalibDevice.WriteString "*cls"
                
        
            Case "5502A"
                PanelForm.STDAction.Caption = CalibratorModel & " Send Command: Clear"
                DoEvents
                CalibDevice.WriteString "*cls"
                
        
            Case "5520A"
                PanelForm.STDAction.Caption = CalibratorModel & " Send Command: Clear"
                DoEvents
                CalibDevice.WriteString "*cls"
                
                
        
            Case "5522A"
                PanelForm.STDAction.Caption = CalibratorModel & " Send Command: Clear"
                DoEvents
                CalibDevice.WriteString "*cls"
                
        
            Case "M3001"
                PanelForm.STDAction.Caption = CalibratorModel & " Send Command: Clear"
                DoEvents
                CalibDevice.WriteString "*cls"
                
        
        End Select
        
    Case "Reset"
    
        Select Case CalibratorModel

            Case "5500A"
                PanelForm.STDAction.Caption = CalibratorModel & ": Send Command: Reset"
                DoEvents
                CalibDevice.WriteString "*RST"
                
                HVImageShow 0, "V"
        
            Case "5502A"
                PanelForm.STDAction.Caption = CalibratorModel & ": Send Command: Reset"
                DoEvents
                CalibDevice.WriteString "*RST"
                
                HVImageShow 0, "V"
        
            Case "5520A"
                PanelForm.STDAction.Caption = CalibratorModel & ": Send Command: Reset"
                DoEvents
                CalibDevice.WriteString "*RST"
                
                HVImageShow 0, "V"
        
            Case "5522A"
                PanelForm.STDAction.Caption = CalibratorModel & ": Send Command: Reset"
                DoEvents
                CalibDevice.WriteString "*RST"
                
                HVImageShow 0, "V"
        
            Case "M3001"
                PanelForm.STDAction.Caption = CalibratorModel & ": Send Command: Reset"
                DoEvents
                CalibDevice.WriteString "*RST"
                
                HVImageShow 0, "V"
        
        End Select
        
        
    Case "Standby"
    
        Select Case CalibratorModel

            Case "5500A"
                PanelForm.STDAction.Caption = CalibratorModel & " Send Command: Standby"
                DoEvents
                Calibrator "Source", "DCV", 0, "V", 0, "Hz", "", 0, 0, ""
                CalibDevice.WriteString "STBY"
                
                
                If ready = 0 Then Exit Sub
                Calibrator "Source", "DCV", 0, "V", 0, "Hz", "", 0, 0, ""
                CalibDevice.WriteString "STBY"
                
        
            Case "5502A"
                PanelForm.STDAction.Caption = CalibratorModel & " Send Command: Standby"
                DoEvents
                Calibrator "Source", "DCV", 0, "V", 0, "Hz", "", 0, 0, ""
                CalibDevice.WriteString "STBY"
                
                
                If ready = 0 Then Exit Sub
                Calibrator "Source", "DCV", 0, "V", 0, "Hz", "", 0, 0, ""
                CalibDevice.WriteString "STBY"
                
        
            Case "5520A"
                PanelForm.STDAction.Caption = CalibratorModel & " Send Command: Standby"
                DoEvents
                Calibrator "Source", "DCV", 0, "V", 0, "Hz", "", 0, 0, ""
                CalibDevice.WriteString "STBY" '"
                
                
                
        
            Case "5522A"
                PanelForm.STDAction.Caption = CalibratorModel & " Send Command: Standby"
                DoEvents
                Calibrator "Source", "DCV", 0, "V", 0, "Hz", "", 0, 0, ""
                CalibDevice.WriteString "STBY"
                
                
                If ready = 0 Then Exit Sub
                Calibrator "Source", "DCV", 0, "V", 0, "Hz", "", 0, 0, ""
                CalibDevice.WriteString "STBY"
                
        
            Case "M3001"
                PanelForm.STDAction.Caption = CalibratorModel & " Send Command: Standby"
                DoEvents
                Calibrator "Source", "DCV", 0, "V", 0, "Hz", "", 0, 0, ""
                CalibDevice.WriteString "STBY"
                
                
                If ready = 0 Then Exit Sub
                Calibrator "Source", "DCV", 0, "V", 0, "Hz", "", 0, 0, ""
                CalibDevice.WriteString "STBY"
                
                
        End Select
        
        
    
    
    
End Select
    
    'Else
    

    
        
       ' instrument.WriteString "*cls"
        
       ' End If
End Sub
