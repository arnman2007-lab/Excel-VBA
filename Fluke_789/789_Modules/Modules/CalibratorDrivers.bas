Attribute VB_Name = "CalibratorDrivers"
'Sub Calibrator(Mode As String, CalFunc As String, CalibParam As Double, CalibParamUnit As String, CalibHertz As Double, CalibHertzUnit As String, Wave As String, OffSet As Double, Duty As Double, ZComp As String)
Sub Calibrator(Mode As String, CalFunc As String, ParamV As Double, ParamUnit As String, ParamHertz As Double, ParamHertzUnit As String, Wave As String, OffSet As Double, Duty As Double, ZComp As String)

Dim OpDone As Integer
CalibMode = Mode
CalibCalFunc = CalFunc
CalibWave = Wave
CalibOffSet = OffSet
CalibDuty = Duty
CalibZComp = ZComp
CalibParam = ParamV
CalibParamUnit = ParamUnit
CalibHertz = ParamHertz
CalibHertzUnit = ParamHertzUnit

CalibratorModel = wsInfo.Range("M9").Value
CalibratorGPIB = wsInfo.Range("M11").Value
CalibratorScopeOption = wsInfo.Range("M12").Value
DMMModel = wsInfo.Range("P9").Value
DMMGPIB = wsInfo.Range("P11").Value
CounterModel = wsInfo.Range("M16").Value
CounterGPIB = wsInfo.Range("M18").Value

'If CalibParam >= 100 And OffValueU = "V" Or OffValueV <= -100 And OffValueU = "V" Then
'MsgBox CalibParam & " " & CalibParamUnit
'If CalibParamUnit = "V" Then
If CalibParam >= 100 And CalibParamUnit = "V" Or CalibParam <= -100 And CalibParamUnit = "V" Then
'MsgBox "here"
HVImageShow CalibParam, CalibParamUnit
End If
                            PanelForm.STDAction.Caption = CalibratorModel & ": Sourcing " & CalibParam & " " & CalibParamUnit & ", " & CalibHertz & " " & CalibHertzUnit
                            DoEvents   ' Forces the label to repaint immediately
'Exit sub if no gpib address
If CalibratorGPIB = "" Then Exit Sub


    'Set ioMgr = New VisaComLib.ResourceManager
    Set ioMgr = New VisaComLib.ResourceManager
    Set CalibDevice = New VisaComLib.FormattedIO488
    Set CalibDevice.IO = ioMgr.Open(CalibratorGPIB)

Select Case CalibratorModel

    Case "5500A"
        Select Case Mode
    
            Case "Source"
            Select Case CalFunc
                    Case "DCV"
                        CalibratorSpecs "DCV", CalibParam, CalibParamUnit, 0, ""
                        If CanDoIt = 1 Then
                            HVImageShow CalibParam, CalibParamUnit
                            'Perform Output
                            CalibClearStatus "Clear"
                            PanelForm.STDAction.Caption = CalibratorModel & ": Sourcing " & CalibParam & " " & CalibParamUnit '& ", " & CalibHertz & " " & CalibHertzUnit
                            DoEvents   ' Forces the label to repaint immediately
                            CalibDevice.WriteString "OUT " & CalibParam & " " & CalibParamUnit & "; Oper"
                            
                            
                            'CalibDevice.WriteString "Oper"
                            
                            
                        Else
                        Exit Sub
                        End If
                    
                    Case "ACV"
                        CalibratorSpecs "ACV", CalibParam, CalibParamUnit, CalibHertz, CalibHertzUnit
                        If CanDoIt = 1 Then
                            HVImageShow CalibParam, CalibParamUnit
                            'Perform Output
                            PanelForm.STDAction.Caption = CalibratorModel & ": Sourcing " & CalibParam & " " & CalibParamUnit & ", " & CalibHertz & " " & CalibHertzUnit
                            DoEvents
                            CalibDevice.WriteString "OUT " & OffValueV & " " & OffValueU & ", " & OffValueHz & " " & OffValueHzU & "; OPER"
                            'CalibDevice.WriteString "OUT " & OffValueHz & " " & OffValueHzU & "; OPER"
                            ' Additional calibration settings
                                If Wave <> "" Then
                                    CalibDevice.WriteString "Wave " & Wave
                                End If
                                If OffSet = "" Then
                                    CalibDevice.WriteString "Offset " & OffSet
                                End If
                                If Duty = "" Then
                                    CalibDevice.WriteString "duty " & Duty
                                End If
                            
                        Else
                        Exit Sub
                        End If
                        
                    Case "DCI"
                        CalibratorSpecs "DCI", CalibParam, CalibParamUnit, "", ""
                        If CanDoIt = 1 Then
                            'Perform Output
                            PanelForm.STDAction.Caption = CalibratorModel & ": Sourcing " & CalibParam & " " & CalibParamUnit '& ", " & CalibHertz & " " & CalibHertzUnit
                            DoEvents
                            CalibDevice.WriteString "OUT " & CalibParam & " " & CalibParamUnit & ", " & CalibHertz & " " & CalibHertzUnit & "; OPER"
                            
                        Else
                            Exit Sub
                        End If
                        
                        
                    Case "ACI"
                        CalibratorSpecs "ACI", CalibParam, CalibParamUnit, CalibHertz, CalibHertzUnit
                        If CanDoIt = 1 Then
                            'Perform Output
                            PanelForm.STDAction.Caption = CalibratorModel & ": Sourcing " & CalibParam & " " & CalibParamUnit & ", " & CalibHertz & " " & CalibHertzUnit
                            DoEvents
                            CalibDevice.WriteString "OUT " & CalibParam & " " & CalibParamUnit & ", " & CalibHertz & " " & CalibHertzUnit & "; OPER"
                            'CalibDevice.WriteString "OUT " & OffValueHz & " " & OffValueHzU & "; OPER"
                            ' Additional calibration settings
                                If Wave <> "" Then
                                    CalibDevice.WriteString "Wave " & Wave
                                End If
                                If OffSet = "" Then
                                    CalibDevice.WriteString "Offset " & OffSet
                                End If
                                If Duty = "" Then
                                    CalibDevice.WriteString "duty " & Duty
                                End If
                            
                        Else
                        Exit Sub
                        End If
                        
                    Case "Ohm"
                        
                        CalibratorSpecs "Ohm", CalibParam, CalibParamUnit, 0, ""
                       
                        If CanDoIt = 1 Then
                            'Perform Output
                            CalibClearStatus "Clear"
                             
                            CalibDevice.WriteString "OUT " & CalibParam & " " & CalibParamUnit & "; ZCOMP " & CalibZComp & "; Oper"
                            
                            
                            'CalibDevice.WriteString "Oper"
                            
                            
                        Else
                        Exit Sub
                        End If
                        
                        
                    Case "Cap"
                        
                        CalibratorSpecs "Cap", CalibParam, CalibParamUnit, 0, ""
                       
                        If CanDoIt = 1 Then
                            'Perform Output
                            CalibClearStatus "Clear"
                             
                            CalibDevice.WriteString "OUT " & CalibParam & " " & CalibParamUnit & "; Oper"
                            
                            
                            'CalibDevice.WriteString "Oper"
                            
                            
                        Else
                        Exit Sub
                        End If
                    
                End Select
            
            Case "Measure"
                Select Case CalFunc
                    Case "Temp"
                    
                        CalibratorSpecs "Temp", CalibParam, CalibParamUnit, 0, ""
                        
                        If CanDoIt = 1 Then
                            'CalibClearStatus "Clear"
                            'CalibDevice.WriteString "tc_meas; tc_type" & " " & CalibParamUnit
                            CalibDevice.WriteString "*RST; TC_TYPE" & " " & CalibParamUnit & "; TC_MEAS " & CalibHertzUnit
                            TestSectBak = TestSect
                            TestSect = 1000
                            UForms "MainForm"
                            TestSect = TestSectBak
                            CalibDevice.WriteString "Val?"
                            inst_value = CalibDevice.ReadString()
                            'MsgBox inst_value
                            ActiveCell.Value = inst_value
                        Else
                            Exit Sub
                        End If
                End Select
            
    
        End Select

    Case "5502A"

        Select Case Mode
    
            Case "Source"
            Select Case CalFunc
                    Case "DCV"
                        CalibratorSpecs "DCV", "CalibParam", "CalibParamUnit", "", ""
                        If CanDoIt = 1 Then
                            HVImageShow CalibParam, CalibParamUnit
                            'Perform Output
                            CalibClearStatus "Clear"
                            PanelForm.STDAction.Caption = CalibratorModel & ": Sourcing " & CalibParam & " " & CalibParamUnit '& ", " & CalibHertz & " " & CalibHertzUnit
                            DoEvents   ' Forces the label to repaint immediately
                            CalibDevice.WriteString "OUT " & CalibParam & " " & CalibParamUnit & "; Oper"
                            
                            
                            'CalibDevice.WriteString "Oper"
                            
                            
                        Else
                        Exit Sub
                        End If
                    
                    Case "ACV"
                        CalibratorSpecs "ACV", "CalibParam", "CalibParamUnit", "", ""
                        If CanDoIt = 1 Then
                            HVImageShow CalibParam, CalibParamUnit
                            'Perform Output
                            PanelForm.STDAction.Caption = CalibratorModel & ": Sourcing " & CalibParam & " " & CalibParamUnit & ", " & CalibHertz & " " & CalibHertzUnit
                            DoEvents   ' Forces the label to repaint immediately
                            CalibDevice.WriteString "OUT " & OffValueV & " " & OffValueU & ", " & OffValueHz & " " & OffValueHzU & "; OPER"
                            'CalibDevice.WriteString "OUT " & OffValueHz & " " & OffValueHzU & "; OPER"
                            ' Additional calibration settings
                                If Wave <> "" Then
                                    CalibDevice.WriteString "Wave " & Wave
                                End If
                                If OffSet = "" Then
                                    CalibDevice.WriteString "Offset " & OffSet
                                End If
                                If Duty = "" Then
                                    CalibDevice.WriteString "duty " & Duty
                                End If
                            
                        Else
                        Exit Sub
                        End If
                        
                    Case "Ohm"
                        
                        CalibratorSpecs "Ohm", CalibParam, CalibParamUnit, 0, ""
                       
                        If CanDoIt = 1 Then
                            'Perform Output
                            CalibClearStatus "Clear"
                             
                            CalibDevice.WriteString "OUT " & CalibParam & " " & CalibParamUnit & "; ZCOMP " & CalibZComp & "; Oper"
                            
                            
                            'CalibDevice.WriteString "Oper"
                            
                            
                        Else
                        Exit Sub
                        End If
                        
                        
                    Case "Cap"
                        
                        CalibratorSpecs "Cap", CalibParam, CalibParamUnit, 0, ""
                       
                        If CanDoIt = 1 Then
                            'Perform Output
                            CalibClearStatus "Clear"
                             
                            CalibDevice.WriteString "OUT " & CalibParam & " " & CalibParamUnit & "; Oper"
                            
                            
                            'CalibDevice.WriteString "Oper"
                            
                            
                        Else
                        Exit Sub
                        End If
                    
                End Select
            
            Case "Measure"
                Select Case CalFunc
                    Case "Temp"
                    
                        CalibratorSpecs "Temp", CalibParam, CalibParamUnit, 0, ""
                        
                        If CanDoIt = 1 Then
                            'CalibClearStatus "Clear"
                            'CalibDevice.WriteString "tc_meas; tc_type" & " " & CalibParamUnit
                            CalibDevice.WriteString "*RST; TC_TYPE" & " " & CalibParamUnit & "; TC_MEAS " & CalibHertzUnit
                            TestSectBak = TestSect
                            TestSect = 1000
                            UForms "MainForm"
                            TestSect = TestSectBak
                            CalibDevice.WriteString "Val?"
                            inst_value = CalibDevice.ReadString()
                            'MsgBox inst_value
                            ActiveCell.Value = inst_value
                        Else
                            Exit Sub
                        End If
                End Select
    
        End Select


    Case "5520A"

        Select Case Mode
    
            Case "Source"
            'MsgBox "here"
                Select Case CalFunc
                    Case "DCV"
                        
                        CalibratorSpecs "DCV", CalibParam, CalibParamUnit, 0, ""
                       
                        If CanDoIt = 1 Then
                            
                            HVImageShow CalibParam, CalibParamUnit
                            'Perform Output
                            CalibClearStatus "Clear"
                            PanelForm.STDAction.Caption = CalibratorModel & ": Sourcing " & CalibParam & " " & CalibParamUnit '& ", " & CalibHertz & " " & CalibHertzUnit
                            DoEvents   ' Forces the label to repaint immediately
                            CalibDevice.WriteString "OUT " & CalibParam & " " & CalibParamUnit & ", " & CalibHertz & " " & CalibHertzUnit
                            CalibDevice.WriteString "Oper"

                            
                            
                        Else
                        Exit Sub
                        End If
                    
                    Case "ACV"
                        'Calibrator "Source", "ACV", 1, "V", 10, "kHz", "Square", 0, 0, "none"
                        
                        CalibratorSpecs "ACV", CalibParam, CalibParamUnit, 0, ""
                        
                        If CanDoIt = 1 Then
                            
                            HVImageShow CalibParam, CalibParamUnit
                            'Perform Output
                            PanelForm.STDAction.Caption = CalibratorModel & ": Sourcing " & CalibParam & " " & CalibParamUnit & ", " & CalibHertz & " " & CalibHertzUnit
                            DoEvents   ' Forces the label to repaint immediately
                            
                            CalibDevice.WriteString "OUT " & CalibParam & " " & CalibParamUnit & ", " & CalibHertz & " " & CalibHertzUnit & "; OPER"
                            
                            'CalibDevice.WriteString "OUT " & OffValueHz & " " & OffValueHzU & "; OPER"
                            ' Additional calibration settings
                                If Wave <> "" Then
                                    CalibDevice.WriteString "Wave " & Wave
                                    
                                End If
                                
                                If OffSet > 0 Then
                                
                                    CalibDevice.WriteString "Offset " & OffSet
                                End If
                                If Duty > 0 Then
                                    CalibDevice.WriteString "duty " & Duty
                                End If
                            
                        Else
                        Exit Sub
                        End If
                        
                    Case "Temp"
                        CalibratorSpecs "Temp", CalibParam, CalibParamUnit, 0, ""
                        If CanDoIt = 1 Then
                            'Perform Output
                            CalibClearStatus "Clear"
                            'MsgBox "here" & CalibParam & CalibParamUnit
                            CalibDevice.WriteString "TSENS_TYPE" & " TC;" & "TC_Type " & CalibParamUnit & "; *WAI"
                            CalibDevice.WriteString "OUT " & CalibParam & " cel; Oper"
                            
                            'OUT 100 CEL
                            'CalibDevice.WriteString "Oper"
                            
                            
                        Else
                            Exit Sub
                        End If
                        
                    Case "Ohm"
                        
                        CalibratorSpecs "Ohm", CalibParam, CalibParamUnit, 0, ""
                       
                        If CanDoIt = 1 Then
                            'Perform Output
                            CalibClearStatus "Clear"
                             
                            CalibDevice.WriteString "OUT " & CalibParam & " " & CalibParamUnit & "; ZCOMP " & CalibZComp & "; Oper"
                            
                            
                            'CalibDevice.WriteString "Oper"
                            
                            
                        Else
                        Exit Sub
                        End If
                        
                        
                    Case "Cap"
                        
                        CalibratorSpecs "Cap", CalibParam, CalibParamUnit, 0, ""
                       MsgBox CalibParam & CalibParamUnit
                        If CanDoIt = 1 Then
                            'Perform Output
                            CalibClearStatus "Clear"
                             
                            CalibDevice.WriteString "OUT " & CalibParam & " " & CalibParamUnit & "; Oper"
                            
                            
                            'CalibDevice.WriteString "Oper"
                            
                            
                        Else
                        Exit Sub
                        End If
                        
                        
                    Case "ACI"
                        
                        CalibratorSpecs "ACI", CalibParam, CalibParamUnit, CalibHertz, CalibHertzUnit
                       
                        If CanDoIt = 1 Then
                            'Perform Output
                            CalibClearStatus "Clear"
                             
                            CalibDevice.WriteString "OUT " & CalibParam & " " & CalibParamUnit & ", " & CalibHertz & " " & CalibHertzUnit & "; Oper"
                            
                            
                            'CalibDevice.WriteString "Oper"
                            
                            
                        Else
                        Exit Sub
                        End If
                        
                        
                    Case "DCI"
                        
                        CalibratorSpecs "DCI", CalibParam, CalibParamUnit, 0, "Hz"
                        
                        If CanDoIt = 1 Then
                            'Perform Output
                            CalibClearStatus "Clear"
                             
                            CalibDevice.WriteString "OUT " & CalibParam & CalibParamUnit & ", " & CalibHertz & CalibHertzUnit & "; Oper"
                                                        
                        Else
                        Exit Sub
                        End If
                    
                End Select
            Case "Measure"
                Select Case CalFunc
                    Case "Temp"
                    
                        CalibratorSpecs "Temp", CalibParam, CalibParamUnit, 0, ""
                        
                        If CanDoIt = 1 Then
                            'CalibClearStatus "Clear"
                            'CalibDevice.WriteString "tc_meas; tc_type" & " " & CalibParamUnit
                            CalibDevice.WriteString "*RST; TC_TYPE" & " " & CalibParamUnit & "; TC_MEAS " & CalibHertzUnit
                            TestSectBak = TestSect
                            TestSect = 1000
                            UForms "MainForm"
                            TestSect = TestSectBak
                            CalibDevice.WriteString "Val?"
                            inst_value = CalibDevice.ReadString()
                            'MsgBox inst_value
                            ActiveCell.Value = inst_value
                        Else
                            Exit Sub
                        End If
                End Select

    
    End Select

    Case "5522A"
    
        Select Case Mode
    
            Case "Source"
            Select Case CalFunc
                    Case "DCV"
                        CalibratorSpecs "DCV", "CalibParam", CalibParamUnit, 0, ""
                        If CanDoIt = 1 Then
                            HVImageShow CalibParam, CalibParamUnit
                            'Perform Output
                            CalibClearStatus "Clear"
                            PanelForm.STDAction.Caption = CalibratorModel & ": Sourcing " & CalibParam & " " & CalibParamUnit '& ", " & CalibHertz & " " & CalibHertzUnit
                            DoEvents   ' Forces the label to repaint immediately
                            CalibDevice.WriteString "OUT " & CalibParam & " " & CalibParamUnit & "; Oper"
                            
                            
                            'CalibDevice.WriteString "Oper"
                            
                            
                        Else
                        Exit Sub
                        End If
                    
                    Case "ACV"
                        CalibratorSpecs "ACV", "CalibParam", "CalibParamUnit", "", ""
                        If CanDoIt = 1 Then
                            HVImageShow CalibParam, CalibParamUnit
                            'Perform Output
                            PanelForm.STDAction.Caption = CalibratorModel & ": Sourcing " & CalibParam & " " & CalibParamUnit & ", " & CalibHertz & " " & CalibHertzUnit
                            DoEvents   ' Forces the label to repaint immediately
                            CalibDevice.WriteString "OUT " & OffValueV & " " & OffValueU & ", " & OffValueHz & " " & OffValueHzU & "; OPER"
                            'CalibDevice.WriteString "OUT " & OffValueHz & " " & OffValueHzU & "; OPER"
                            ' Additional calibration settings
                                If Wave <> "" Then
                                    CalibDevice.WriteString "Wave " & Wave
                                End If
                                If OffSet = "" Then
                                    CalibDevice.WriteString "Offset " & OffSet
                                End If
                                If Duty = "" Then
                                    CalibDevice.WriteString "duty " & Duty
                                End If
                            
                        Else
                        Exit Sub
                        End If
                    
                End Select
            
            Case "Measure"
                Select Case CalFunc
                    Case "Temp"
                    
                        CalibratorSpecs "Temp", CalibParam, CalibParamUnit, 0, ""
                        
                        If CanDoIt = 1 Then
                            'CalibClearStatus "Clear"
                            'CalibDevice.WriteString "tc_meas; tc_type" & " " & CalibParamUnit
                            CalibDevice.WriteString "*RST; TC_TYPE" & " " & CalibParamUnit & "; TC_MEAS " & CalibHertzUnit
                            TestSectBak = TestSect
                            TestSect = 1000
                            UForms "MainForm"
                            TestSect = TestSectBak
                            CalibDevice.WriteString "Val?"
                            inst_value = CalibDevice.ReadString()
                            'MsgBox inst_value
                            ActiveCell.Value = inst_value
                        Else
                            Exit Sub
                        End If
                End Select
    
        End Select

    Case "M3001"
    
        Select Case Mode
    
            Case "Source"
            Select Case CalFunc
                    Case "DCV"
                        HVImageShow CalibParam, CalibParamUnit
                        CalibratorSpecs "DCV", "CalibParam", "CalibParamUnit", "", ""
                        If CanDoIt = 1 Then
                            'Perform Output
                            CalibClearStatus "Clear"
                            
                            CalibDevice.WriteString "OUT " & CalibParam & " " & CalibParamUnit & "; Oper"
                            
                            
                            'CalibDevice.WriteString "Oper"
                            
                            
                        Else
                        Exit Sub
                        End If
                    
                    Case "ACV"
                        HVImageShow CalibParam, CalibParamUnit
                        CalibratorSpecs "ACV", "CalibParam", "CalibParamUnit", "", ""
                        If CanDoIt = 1 Then
                        Else
                        Exit Sub
                        End If
                    
                End Select
            
            Case "Measure"
    
        End Select

End Select


End Sub

Sub CalibratorScope(Options As String, volt As Double, VoltUnit As String, Hertz As Double, HertzUnit As String, Wave As String, OffSet As Double, Duty As Double)
'*OPT? returns scope pack if any
Select Case CalibratorModel

    Case "5500A"
        Select Case Options
            
            Select Case CalibratorScopeOptions
            
                Case "SC300"
                
                Case "SC600"
                
                Case "SC1100"
                
            End Select
    
            Case "Volt"
            
            Case "Edge"
            
            Case "Levsine"
            
            Case "Marker"
    
            Case "Pulse"
            
            Case "WaveGen"
            
            Case "Video"
            
            Case "Meas Z"
            
            Case "Overld"
            
    
        End Select

    Case "5502A"

        Select Case Mode
    
            Case "Volt"
            
            Case "Edge"
            
            Case "Levsine"
            
            Case "Marker"
    
            Case "Pulse"
            
            Case "WaveGen"
            
            Case "Video"
            
            Case "Meas Z"
            
            Case "Overld"
    
        End Select


    Case "5520A"

        Select Case Mode
    
            Case "Volt"
            
            Case "Edge"
            
            Case "Levsine"
            
            Case "Marker"
    
            Case "Pulse"
            
            Case "WaveGen"
            
            Case "Video"
            
            Case "Meas Z"
            
            Case "Overld"
    
        End Select

    Case "5522A"
    
        Select Case Mode
    
            Case "Volt"
            
            Case "Edge"
            
            Case "Levsine"
            
            Case "Marker"
    
            Case "Pulse"
            
            Case "WaveGen"
            
            Case "Video"
            
            Case "Meas Z"
            
            Case "Overld"
    
        End Select

    Case "M3001"
    
        Select Case Mode
    
            Case "Source"
            
            Case "Measure"
    
        End Select

End Select


End Sub


