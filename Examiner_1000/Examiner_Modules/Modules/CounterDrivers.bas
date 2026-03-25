Attribute VB_Name = "CounterDrivers"
Sub Counter(Mode As String, CalFunc As String, Channel As String, CalArg2 As String)
'Counter   "Measure",        "Freq",             "1",                ""

CounterMode = Mode
CounterCalFunc = CalFunc
CounterParam = Channel
CounterParamUnit = CalArg2


CalibratorModel = wsInfo.Range("M9").Value
CalibratorGPIB = wsInfo.Range("M11").Value
CalibratorScopeOption = wsInfo.Range("M12").Value
DMMModel = wsInfo.Range("P9").Value
DMMGPIB = wsInfo.Range("P11").Value
CounterModel = wsInfo.Range("M16").Value
CounterGPIB = wsInfo.Range("M18").Value
    
'Exit sub if no gpib address
If CounterGPIB = "" Then Exit Sub
    Set ioMgr = New VisaComLib.ResourceManager
    Set CounterDevice = New VisaComLib.FormattedIO488
    Set CounterDevice.IO = ioMgr.Open(CounterGPIB)

Select Case CounterModel

    Case "PM6681"
        Select Case Mode
    
            Case "Source"
            
            Case "Measure"
                Select Case CalFunc
                    Case "Freq"
                        'CounterDevice.WriteString "*RST"
                        CounterDevice.WriteString ":FUNC '" & CalFunc & "'"
                        CounterDevice.WriteString ":CONF:Freq (@1);:INP:FILT ON;:READ?"
                        OpDone = CounterDevice.ReadString()
                        ActiveCell.Value = OpDone
                End Select
    
        End Select

    Case "5502A"

        Select Case Mode
    
            Case "Source"
            
            Case "Measure"
    
        End Select


    Case "5520A"

        Select Case Mode
    
            Case "Source"
                Select Case CalFunc
                    Case "DCV"
                        CheckSpecs "DCV"
                        If CanDoIt = 1 Then
                            'Perform Output
                            MsgBox CalibParam & CalibParamUnit & CalibHertz & CalibHertzUnit
                            CalibDevice.WriteString "OUT " & CalibParam & " " & CalibParamUnit & ", " & CalibHertz & " " & CalibHertzUnit & "; OPER; *OPC?"
                            OpDone = CalibDevice.ReadString()
                        Else
                        Exit Sub
                        End If
                    
                    Case "ACV"
                        CheckSpecs "ACV"
                        If CanDoIt = 1 Then
                            'Perform Output
                            MsgBox OffValueV & OffValueU & OffValueHz & OffValueHzU & Wave & OffSet & Duty
                            ' Send output string to device
                            CalibDevice.WriteString "OUT " & OffValueV & " " & OffValueU & ", " & OffValueHz & " " & OffValueHzU & "; OPER; *OPC?"
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
    
    End Select

    Case "5522A"
    
        Select Case Mode
    
            Case "Source"
            
            Case "Measure"
    
        End Select

    Case "M3001"
    
        Select Case Mode
    
            Case "Source"
            
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


Sub DMMs(DMMFunc As String)

Select Case DMMModel

Case "3458A"

Case "8508A"

Case "34401A"

End Select

End Sub

Sub Counters(CounterFunc As String)

Select Case CounterModel

Case "PM6680"

Case "PM6681"

End Select

End Sub


