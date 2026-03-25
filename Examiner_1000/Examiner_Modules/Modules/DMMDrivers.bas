Attribute VB_Name = "DMMDrivers"


Sub DMM(CalFunc As String, CalArg As String)


CalibratorModel = wsInfo.Range("M9").Value
CalibratorGPIB = wsInfo.Range("M11").Value
CalibratorScopeOption = wsInfo.Range("M12").Value
DMMModel = wsInfo.Range("P9").Value
DMMGPIB = wsInfo.Range("P11").Value
CounterModel = wsInfo.Range("M16").Value
CounterGPIB = wsInfo.Range("M18").Value

PanelForm.STDAction.Caption = DMMModel & " Send Command: " & CalFunc & " " & CalArg
                DoEvents

'Exit sub if no gpib address
If DMMGPIB = "" Then Exit Sub



    Set ioMgr = New VisaComLib.ResourceManager
    Set DMMDevice = New VisaComLib.FormattedIO488
    Set DMMDevice.IO = ioMgr.Open(DMMGPIB)


DMMDevice.WriteString "END ALWAYS"
     
Select Case DMMModel

    Case "3458A"
        Select Case CalFunc
    
            Case "NPLC"
                PanelForm.STDAction.Caption = DMMModel & " Send Command: NPLC " & CalArg
                DoEvents
                DMMDevice.WriteString "NPLC " & CalArg
                            
            
            Case "MMath"
                PanelForm.STDAction.Caption = DMMModel & " Send Command: MMath " & CalArg
                DoEvents
                DMMDevice.WriteString "MMath " & CalArg
                            
                        
            Case "NRDGS"
                PanelForm.STDAction.Caption = DMMModel & " Send Command: NRDGS " & CalArg
                DoEvents
                DMMDevice.WriteString "NRDGS " & CalArg
                             
                        
            Case "Func"
                PanelForm.STDAction.Caption = DMMModel & " Send Command: Func " & CalArg
                DoEvents
                DMMDevice.WriteString "Func " & CalArg
                             
                        
            Case "Range"
                PanelForm.STDAction.Caption = DMMModel & " Send Command: Range " & CalArg
                DoEvents
                DMMDevice.WriteString "Range " & CalArg
                             
                        
            Case "TRIG"
                PanelForm.STDAction.Caption = DMMModel & " Send Command: Trig " & CalArg
                DoEvents
                Application.Wait (Now + TimeValue("0:00:2"))
                DMMDevice.WriteString "TRIG " & CalArg
                DMMQuery = DMMDevice.ReadString()
                FixReading (DMMQuery)
                
                
                
                             
                        
            Case "RMATH"
                PanelForm.STDAction.Caption = DMMModel & " Send Command: RMath " & CalArg
                DoEvents
                DMMDevice.WriteString "RMATH " & CalArg
                DMMQuery = DMMDevice.ReadString()
                DMMQuery = Replace(Replace(DMMQuery, vbCr, ""), vbLf, "")
                ActiveCell.Value = DMMQuery
                
                
            
            Case "END"
                    PanelForm.STDAction.Caption = DMMModel & " Send Command: END " & CalArg
                DoEvents
                DMMSpecs CalFunc, "", "", ""
                DMMDevice.WriteString "END " & CalArg
                        
            Case "DELAY"
                    PanelForm.STDAction.Caption = DMMModel & " Send Command: DELAY " & CalArg
                DoEvents
                DMMDevice.WriteString "DELAY " & CalArg
            
            Case "MATH"
                PanelForm.STDAction.Caption = DMMModel & " Send Command: Math " & CalArg
                DoEvents
                DMMDevice.WriteString "MATH " & CalArg
                
            Case "RESET"
                PanelForm.STDAction.Caption = DMMModel & " Send Command: Reset"
                DoEvents
                DMMDevice.WriteString "RESET"
        End Select

    'Case New Case Goes here
    
    
    



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



