Attribute VB_Name = "CalibratorSpecifications"

Sub CalibratorSpecs(CalFunc As String, CalibParam As Double, CalibParamUnit As String, CalibHertz As Double, CalibHertzUnit As String)
   
    'CalibratorSpecs "DCV", CalibParam, "CalibParamUnit", 0, ""
    'Select Case CalibParamUnit
    '   Case "uV"
    '  Case "mV"
    ' Case "V"
    'Case "kV"
    'End Select
    'MsgBox "here"
    With Worksheets("Information")
        CalibratorModel = .Range("M9").Value
        DMMModel = .Range("P9").Value
        CounterModel = .Range("M16").Value
    End With
     
   ' MsgBox CalFunc & CalibParam & CalibParamUnit & CalibratorModel & CanDoIt
    Select Case CalibratorModel
        
        Case "5520A"
            
            Select Case CalFunc
                
                Case "DCV"
                    '------------------------Begin DCV Check------------------------------
                    
                    Select Case CalibParamUnit
                        Case "uV"
                            If CalibParam >= -999.9999 And CalibParam <= 999.9999 Then
                                CanDoIt = 1
                            Else
                                
                                CalibratorSpecs "DCV", CalibParam / 1000, "mV", "", ""
                            End If
                            
                        Case "mV"
                            If CalibParam >= -999.9999 And CalibParam <= 999.9999 Then
                                CanDoIt = 1
                            Else
                                
                                CalibratorSpecs "DCV", CalibParam / 1000, "V", "", ""
                            End If
                        Case "V"
                            
                            If CalibParam >= -1000 And CalibParam <= 1000 Then
                                CanDoIt = 1
                            Else
                                CalibratorSpecs "DCV", CalibParam / 1000, "kV", "", ""
                                
                            End If
                            
                        Case "kV"
                            
                            If CalibParam >= -1 And CalibParam <= 1 Then
                                CanDoIt = 1
                                
                            Else
                                MsgBox "Calibrator Is unable To Source " & CalibParam & CalibParamUnit
                                CanDoIt = 0
                            End If
                            
                    End Select
                    '------------------------End DCV Check--------------------------------
                    
                Case "ACV"
                    '------------------------Begin ACV Check------------------------------
                    
                    Select Case CalibParamUnit
                        Case "mV"
                            If CalibParam >= 1 And CalibParam <= 329.999 Then         '-----------------Check to see if Calibrator can go lower voltage than 1 .1,.001
                                Select Case CalibHertzUnit
                                    Case "Hz"
                                        If CalibHertz >= 10 And CalibHertz <= 500000 Then
                                            CanDoIt = 1
                                        Else
                                            CalibratorSpecs "ACV", CalibParam, "mV", CalibHertz / 1000, "kHz"
                                        End If
                                        
                                    Case "kHz"
                                        If CalibHertz >= 0.01 And CalibHertz <= 500 Then
                                            CanDoIt = 1
                                        Else
                                            MsgBox "Calibrator Is unable To Source " & CalibParam & CalibParamUnit & " at " & CalibHertz & CalibHertzUnit
                                        End If
                                        
                                End Select
                                CanDoIt = 1
                            Else
                                
                                CalibratorSpecs "ACV", CalibParam / 1000, "V", CalibHertz, CalibHertzUnit
                            End If
                            
                        Case "V"
                            If CalibParam >= 0.001 And CalibParam <= 3.2999 Then
                                Select Case CalibHertzUnit
                                    Case "Hz"
                                        If CalibHertz >= 10 And CalibHertz <= 500000 Then
                                            CanDoIt = 1
                                        Else
                                            CalibratorSpecs "ACV", CalibParam, "V", CalibHertz / 1000, "kHz"
                                        End If
                                        
                                    Case "kHz"
                                        If CalibHertz >= 0.01 And CalibHertz <= 500 Then
                                            CanDoIt = 1
                                        Else
                                            MsgBox "Calibrator Is unable To Source " & CalibParam & CalibParamUnit & " at " & CalibHertz & CalibHertzUnit
                                        End If
                                        
                                End Select
                                CanDoIt = 1
                                
                            ElseIf CalibParam >= 3.3 And CalibParam <= 32.9999 Then
                                Select Case CalibHertzUnit
                                    Case "Hz"
                                        If CalibHertz >= 10 And CalibHertz <= 100000 Then
                                            CanDoIt = 1
                                        Else
                                            CalibratorSpecs "ACV", CalibParam, "V", CalibHertz / 1000, "kHz"
                                        End If
                                        
                                    Case "kHz"
                                        If CalibHertz >= 0.01 And CalibHertz <= 100 Then
                                            CanDoIt = 1
                                        Else
                                            MsgBox "Calibrator Is unable To Source " & CalibParam & CalibParamUnit & " at " & CalibHertz & CalibHertzUnit
                                        End If
                                        
                                End Select
                                CanDoIt = 1
                                
                            ElseIf CalibParam >= 33 And CalibParam <= 329.9999 Then
                                Select Case CalibHertzUnit
                                    Case "Hz"
                                        If CalibHertz >= 45 And CalibHertz <= 100000 Then
                                            CanDoIt = 1
                                        Else
                                            CalibratorSpecs "ACV", CalibParam, "V", CalibHertz / 1000, "kHz"
                                        End If
                                        
                                    Case "kHz"
                                        If CalibHertz >= 0.045 And CalibHertz <= 100 Then
                                            CanDoIt = 1
                                        Else
                                            MsgBox "Calibrator Is unable To Source " & CalibParam & CalibParamUnit & " at " & CalibHertz & CalibHertzUnit
                                        End If
                                        
                                End Select
                                CanDoIt = 1
                                
                            ElseIf CalibParam >= 330 And CalibParam <= 1020 Then
                                Select Case CalibHertzUnit
                                    Case "Hz"
                                        If CalibHertz >= 45 And CalibHertz <= 10000 Then
                                            CanDoIt = 1
                                        Else
                                            CalibratorSpecs "ACV", CalibParam, "V", CalibHertz / 1000, "kHz"
                                        End If
                                        
                                    Case "kHz"
                                        If CalibHertz >= 0.045 And CalibHertz <= 10 Then
                                            CanDoIt = 1
                                        Else
                                            MsgBox "Calibrator Is unable To Source " & CalibParam & CalibParamUnit & " at " & CalibHertz & CalibHertzUnit
                                        End If
                                        
                                End Select
                                CanDoIt = 1
                            
                            
                            Else
                                
                                CalibratorSpecs "ACV", CalibParam / 1000, "kV", CalibHertz, CalibHertzUnit
                            End If

                    Case "kV"
                            
                            If CalibParam >= 0.000001 And CalibParam <= 1.02 Then
                                CanDoIt = 1
                                
                            Else
                                MsgBox "Calibrator Is unable To Source " & CalibParam & CalibParamUnit & " at " & CalibHertz & CalibHertzUnit
                                CanDoIt = 0
                            End If
                            
                    End Select
'------------------------End ACV Check--------------------------------
                    
'------------------------Begin DCA Check------------------------------
                Case "DCI"
                    
                    Select Case CalibParamUnit
                        Case "uA"
                            If CalibParam >= -999.9999 And CalibParam <= 999.9999 Then
                                CanDoIt = 1
                                
                            Else
                                
                                CalibratorSpecs "DCA", CalibParam / 1000, "mA", "", ""
                            End If
                            
                        Case "mA"
                            If CalibParam >= -999.9999 And CalibParam <= 999.9999 Then
                                CanDoIt = 1
                                
                            Else
                                
                                CalibratorSpecs "DCA", CalibParam / 1000, "A", "", ""
                            End If
                            
                        Case "A"
                            
                            If CalibParam >= -20.5 And CalibParam <= 20.5 Then
                                CanDoIt = 1
                                
                            Else
                                MsgBox "Calibrator Is unable To Source " & CalibParam & CalibParamUnit
                                CanDoIt = 0
                            End If
                            
                        Case "kA"
                            
                            MsgBox "Calibrator Is unable To Source " & CalibParam & CalibParamUnit
                                                        
                        Case "MA"
                            
                            MsgBox "Calibrator Is unable To Source " & CalibParam & CalibParamUnit
                            
'------------------------End DCA Check--------------------------------
                    End Select
                    
'------------------------Begin ACI Check------------------------------
                Case "ACI"
                    
                    Select Case CalibParamUnit
                        Case "uA"
                            If CalibParam >= 1 And CalibParam <= 329.999 Then         '-----------------Check to see if Calibrator can go lower voltage than 1 .1,.001
                                Select Case CalibHertzUnit
                                    Case "Hz"
                                        If CalibHertz >= 10 And CalibHertz <= 30000 Then
                                            CanDoIt = 1
                                        Else
                                            CalibratorSpecs "ACI", CalibParam, "uA", CalibHertz / 1000, "kHz"
                                        End If
                                        
                                    Case "kHz"
                                        If CalibHertz >= 0.01 And CalibHertz <= 30 Then
                                            CanDoIt = 1
                                        Else
                                            MsgBox "Calibrator Is unable To Source " & CalibParam & CalibParamUnit & " at " & CalibHertz & CalibHertzUnit
                                        End If
                                        
                                End Select
                                CanDoIt = 1
                            Else
                                
                                CalibratorSpecs "ACI", CalibParam / 1000, "mA", CalibHertz, CalibHertzUnit
                            End If
                            
                        Case "mA"
                            If CalibParam >= 0.001 And CalibParam <= 3.2999 Then
                                Select Case CalibHertzUnit
                                    Case "Hz"
                                        If CalibHertz >= 10 And CalibHertz <= 30000 Then
                                            CanDoIt = 1
                                        Else
                                            CalibratorSpecs "ACI", CalibParam, "mA", CalibHertz / 1000, "kHz"
                                        End If
                                        
                                    Case "kHz"
                                        If CalibHertz >= 0.01 And CalibHertz <= 30 Then
                                            CanDoIt = 1
                                        Else
                                            MsgBox "Calibrator Is unable To Source " & CalibParam & CalibParamUnit & " at " & CalibHertz & CalibHertzUnit
                                        End If
                                        
                                End Select
                                CanDoIt = 1
                                
                            ElseIf CalibParam >= 3.3 And CalibParam <= 32.9999 Then
                                Select Case CalibHertzUnit
                                    Case "Hz"
                                        If CalibHertz >= 10 And CalibHertz <= 30000 Then
                                            CanDoIt = 1
                                        Else
                                            CalibratorSpecs "ACI", CalibParam, "mA", CalibHertz / 1000, "kHz"
                                        End If
                                        
                                    Case "kHz"
                                        If CalibHertz >= 0.01 And CalibHertz <= 30 Then
                                            CanDoIt = 1
                                        Else
                                            MsgBox "Calibrator Is unable To Source " & CalibParam & CalibParamUnit & " at " & CalibHertz & CalibHertzUnit
                                        End If
                                        
                                End Select
                                CanDoIt = 1
                                
                            ElseIf CalibParam >= 33 And CalibParam <= 329.9999 Then
                                Select Case CalibHertzUnit
                                    Case "Hz"
                                        If CalibHertz >= 10 And CalibHertz <= 30000 Then
                                            CanDoIt = 1
                                        Else
                                            CalibratorSpecs "ACI", CalibParam, "mA", CalibHertz / 1000, "kHz"
                                        End If
                                        
                                    Case "kHz"
                                        If CalibHertz >= 0.01 And CalibHertz <= 30 Then
                                            CanDoIt = 1
                                        Else
                                            MsgBox "Calibrator Is unable To Source " & CalibParam & CalibParamUnit & " at " & CalibHertz & CalibHertzUnit
                                        End If
                                        
                                End Select
                                CanDoIt = 1
                                
                            ElseIf CalibParam >= 330 And CalibParam <= 1099.99 Then
                                Select Case CalibHertzUnit
                                    Case "Hz"
                                        If CalibHertz >= 10 And CalibHertz <= 10000 Then
                                            CanDoIt = 1
                                        Else
                                            CalibratorSpecs "ACI", CalibParam, "mA", CalibHertz / 1000, "kHz"
                                        End If
                                        
                                    Case "kHz"
                                        If CalibHertz >= 0.01 And CalibHertz <= 10 Then
                                            CanDoIt = 1
                                        Else
                                            MsgBox "Calibrator Is unable To Source " & CalibParam & CalibParamUnit & " at " & CalibHertz & CalibHertzUnit
                                        End If
                                        
                                End Select
                                CanDoIt = 1
                            
                            
                            Else
                                
                                CalibratorSpecs "ACI", CalibParam / 1000, "kV", CalibHertz, CalibHertzUnit
                            End If

                    Case "A"
                            
                            If CalibParam >= 1 And CalibParam <= 2.99999 Then
                                
                                Select Case CalibHertzUnit
                                    Case "Hz"
                                        If CalibHertz >= 10 And CalibHertz <= 10000 Then
                                            CanDoIt = 1
                                        Else
                                            CalibratorSpecs "ACI", CalibParam, "A", CalibHertz / 1000, "kHz"
                                        End If
                                        
                                    Case "kHz"
                                        If CalibHertz >= 0.01 And CalibHertz <= 10 Then
                                            CanDoIt = 1
                                        Else
                                            MsgBox "Calibrator Is unable To Source " & CalibParam & CalibParamUnit & " at " & CalibHertz & CalibHertzUnit
                                        End If
                                        
                                End Select
                                CanDoIt = 1
                                
                            ElseIf CalibParam >= 3 And CalibParam <= 10.9999 Then
                                Select Case CalibHertzUnit
                                    Case "Hz"
                                        If CalibHertz >= 45 And CalibHertz <= 5000 Then
                                            CanDoIt = 1
                                        Else
                                            CalibratorSpecs "ACI", CalibParam, "A", CalibHertz / 1000, "kHz"
                                        End If
                                        
                                    Case "kHz"
                                        If CalibHertz >= 0.045 And CalibHertz <= 5 Then
                                            CanDoIt = 1
                                        Else
                                            MsgBox "Calibrator Is unable To Source " & CalibParam & CalibParamUnit & " at " & CalibHertz & CalibHertzUnit
                                        End If
                                        
                                End Select
                                CanDoIt = 1
                                
                            ElseIf CalibParam >= 11 And CalibParam <= 20.5 Then
                                Select Case CalibHertzUnit
                                    Case "Hz"
                                        If CalibHertz >= 45 And CalibHertz <= 5000 Then
                                            CanDoIt = 1
                                        Else
                                            CalibratorSpecs "ACI", CalibParam, "A", CalibHertz / 1000, "kHz"
                                        End If
                                        
                                    Case "kHz"
                                        If CalibHertz >= 0.045 And CalibHertz <= 5 Then
                                            CanDoIt = 1
                                        Else
                                            MsgBox "Calibrator Is unable To Source " & CalibParam & CalibParamUnit & " at " & CalibHertz & CalibHertzUnit
                                        End If
                                        
                                End Select
                                CanDoIt = 1
                                
                            Else
                                MsgBox "Calibrator Is unable To Source " & CalibParam & CalibParamUnit & " at " & CalibHertz & CalibHertzUnit
                                CanDoIt = 0
                            End If
                            
                    End Select
'------------------------End ACI Check--------------------------------
                    

                    
                Case "Temp"
'------------------------Begin Temp Check in C Degs------------------------------
                    
                    Select Case CalibParamUnit
                        Case "B"
                            If CalibParam >= 600 And CalibParam <= 1820 Then
                                CanDoIt = 1
                            Else
                                MsgBox "Calibrator Is unable To Source " & CalibParam & CalibParamUnit
                                CanDoIt = 0
                            End If
                            
                        Case "C"
                            If CalibParam >= 0 And CalibParam <= 2316 Then
                                CanDoIt = 1
                            Else
                                MsgBox "Calibrator Is unable To Source " & CalibParam & CalibParamUnit
                                CanDoIt = 0
                            End If
                        Case "E"
                            If CalibParam >= -250 And CalibParam <= 1000 Then
                                CanDoIt = 1
                            Else
                                MsgBox "Calibrator Is unable To Source " & CalibParam & CalibParamUnit
                                CanDoIt = 0
                            End If
                            
                        Case "J"
                            If CalibParam >= -210 And CalibParam <= 1200 Then
                                CanDoIt = 1
                            Else
                                MsgBox "Calibrator Is unable To Source " & CalibParam & CalibParamUnit
                                CanDoIt = 0
                            End If
                            
                        Case "K"
                            If CalibParam >= -200 And CalibParam <= 1372 Then
                                CanDoIt = 1
                            Else
                                MsgBox "Calibrator Is unable To Source " & CalibParam & CalibParamUnit
                                CanDoIt = 0
                            End If
                            
                        Case "L"
                            If CalibParam >= -200 And CalibParam <= 900 Then
                                CanDoIt = 1
                            Else
                                MsgBox "Calibrator Is unable To Source " & CalibParam & CalibParamUnit
                                CanDoIt = 0
                            End If
                            
                        Case "N"
                            If CalibParam >= -200 And CalibParam <= 1300 Then
                                CanDoIt = 1
                            Else
                                MsgBox "Calibrator Is unable To Source " & CalibParam & CalibParamUnit
                                CanDoIt = 0
                            End If
                            
                        Case "R"
                            If CalibParam >= 0 And CalibParam <= 1767 Then
                                CanDoIt = 1
                            Else
                                MsgBox "Calibrator Is unable To Source " & CalibParam & CalibParamUnit
                                CanDoIt = 0
                            End If
                            
                        Case "S"
                            If CalibParam >= 0 And CalibParam <= 1767 Then
                                CanDoIt = 1
                            Else
                                MsgBox "Calibrator Is unable To Source " & CalibParam & CalibParamUnit
                                CanDoIt = 0
                            End If
                            
                        Case "T"
                            If CalibParam >= -250 And CalibParam <= 400 Then
                                CanDoIt = 1
                            Else
                                MsgBox "Calibrator Is unable To Source " & CalibParam & CalibParamUnit
                                CanDoIt = 0
                            End If
                            
                        Case "U"
                            If CalibParam >= -200 And CalibParam <= 600 Then
                                CanDoIt = 1
                            Else
                                MsgBox "Calibrator Is unable To Source " & CalibParam & CalibParamUnit
                                CanDoIt = 0
                            End If
                            
                    End Select
'------------------------End Temp Check in C Degs--------------------------------
                    
                    '------------------------Begin Ohm Check------------------------------
                Case "Ohm"
                    
                    Select Case CalibParamUnit
                        Case "uOhm"
                             'MsgBox "uOhm"
                            If CalibParam >= 0 And CalibParam <= 999.9999 Then
                                CanDoIt = 1
                                
                            Else
                                
                                CalibratorSpecs "Ohm", CalibParam / 1000, "mOhm", "", ""
                            End If
                            
                        Case "mOhm"
                         'MsgBox "mOhm"
                            If CalibParam >= 0 And CalibParam <= 999.9999 Then
                                CanDoIt = 1
                                
                            Else
                                
                                CalibratorSpecs "Ohm", CalibParam / 1000, "Ohm", "", ""
                            End If
                            
                        Case "Ohm"
                             'MsgBox "Ohm"
                            If CalibParam >= 0 And CalibParam <= 999.9999 Then
                                CanDoIt = 1
                                
                            Else
                                CalibratorSpecs "Ohm", CalibParam / 1000, "kOhm", "", ""
                            End If
                            
                        Case "kOhm"
                            ' MsgBox "kOhm"
                            If CalibParam >= 0 And CalibParam <= 999.9999 Then
                                CanDoIt = 1
                               
                            Else
                                CalibratorSpecs "Ohm", CalibParam / 1000, "MOhm", "", ""
                                CanDoIt = 0
                            End If
                            
                        Case "MOhm"
                            'MsgBox "MOhm"
                            If CalibParam >= 0 And CalibParam <= 999.9999 Then
                                CanDoIt = 1
                                
                            Else
                                CalibratorSpecs "Ohm", CalibParam / 1000, "GOhm", "", ""
                                CanDoIt = 0
                            End If
                            
                        Case "GOhm"
                            
                            If CalibParam >= 0 And CalibParam <= 1.1 Then
                                CanDoIt = 1
                                
                            Else
                                MsgBox "Calibrator Is unable To Source " & CalibParam & CalibParamUnit
                                CanDoIt = 0
                            End If
                            
                            '------------------------End Ohm Check--------------------------------
                    End Select
                
                Case "Cap"
'------------------------Begin Cap Check------------------------------
                   
                    Select Case "CalibParamUnit"
                        Case "pF"
                        MsgBox CalibParam & CalibParamUnit & "2"
                            If CalibParam >= 190 And CalibParam <= 999.9999 Then
                                CanDoIt = 1
                            Else
                                
                                CalibratorSpecs "Cap", CalibParam / 1000, "nF", "", ""
                            End If
                            
                        Case "nF"
                            If CalibParam >= 1 And CalibParam <= 999.9999 Then
                                CanDoIt = 1
                            Else
                                
                                CalibratorSpecs "Cap", CalibParam / 1000, "uF", "", ""
                            End If
                            
                        Case "uF"
                            
                            If CalibParam >= 1 And CalibParam <= 999.999 Then
                                CanDoIt = 1
                            Else
                                CalibratorSpecs "Cap", CalibParam / 1000, "mF", "", ""
                                
                            End If
                            
                        Case "mF"
                            
                            If CalibParam >= 1 And CalibParam <= 110 Then
                                CanDoIt = 1
                                
                            Else
                                MsgBox "Calibrator Is unable To Source " & CalibParam & CalibParamUnit
                                CanDoIt = 0
                            End If
                            
                    End Select
'------------------------End Cap Check--------------------------------
                    
            End Select
            
        Case "5500A"
            Select Case CalFunc
                
                Case "DCV"
                    '------------------------Begin DCV Check------------------------------
                    
                    Select Case CalibParamUnit
                        Case "uV"
                            If CalibParam >= -999.9999 And CalibParam <= 999.9999 Then
                                CanDoIt = 1
                            Else
                                
                                CalibratorSpecs "DCV", CalibParam / 1000, "mV", "", ""
                            End If
                            
                        Case "mV"
                            If CalibParam >= -999.9999 And CalibParam <= 999.9999 Then
                                CanDoIt = 1
                            Else
                                
                                CalibratorSpecs "DCV", CalibParam / 1000, "V", "", ""
                            End If
                        Case "V"
                            
                            If CalibParam >= -1000 And CalibParam <= 1000 Then
                                CanDoIt = 1
                            Else
                                CalibratorSpecs "DCV", CalibParam / 1000, "kV", "", ""
                            End If
                            
                        Case "kV"
                            
                            If CalibParam >= -1 And CalibParam <= 1 Then
                                CanDoIt = 1
                                
                            Else
                                MsgBox "Calibrator Is unable To Source " & CalibParam & CalibParamUnit
                                CanDoIt = 0
                            End If
                            '------------------------End DCV Check--------------------------------
                    End Select
                    
                Case "ACV"
'------------------------Begin ACV Check------------------------------
                    
                    Select Case CalibParamUnit
                        Case "mV"
                            If CalibParam >= 1 And CalibParam <= 329.999 Then         '-----------------Check to see if Calibrator can go lower voltage than 1 .1,.001
                                Select Case CalibHertzUnit
                                    Case "Hz"
                                        If CalibHertz >= 10 And CalibHertz <= 500000 Then
                                            CanDoIt = 1
                                        Else
                                            CalibratorSpecs "ACV", CalibParam, "mV", CalibHertz / 1000, "kHz"
                                        End If
                                        
                                    Case "kHz"
                                        If CalibHertz >= 0.01 And CalibHertz <= 500 Then
                                            CanDoIt = 1
                                        Else
                                            MsgBox "Calibrator Is unable To Source " & CalibParam & CalibParamUnit & " at " & CalibHertz & CalibHertzUnit
                                        End If
                                        
                                End Select
                                CanDoIt = 1
                            Else
                                
                                CalibratorSpecs "ACV", CalibParam / 1000, "V", CalibHertz, CalibHertzUnit
                            End If
                            
                        Case "V"
                            If CalibParam >= 0.001 And CalibParam <= 3.2999 Then
                                Select Case CalibHertzUnit
                                    Case "Hz"
                                        If CalibHertz >= 10 And CalibHertz <= 500000 Then
                                            CanDoIt = 1
                                        Else
                                            CalibratorSpecs "ACV", CalibParam, "V", CalibHertz / 1000, "kHz"
                                        End If
                                        
                                    Case "kHz"
                                        If CalibHertz >= 0.01 And CalibHertz <= 500 Then
                                            CanDoIt = 1
                                        Else
                                            MsgBox "Calibrator Is unable To Source " & CalibParam & CalibParamUnit & " at " & CalibHertz & CalibHertzUnit
                                        End If
                                        
                                End Select
                                CanDoIt = 1
                                
                            ElseIf CalibParam >= 3.3 And CalibParam <= 32.9999 Then
                                Select Case CalibHertzUnit
                                    Case "Hz"
                                        If CalibHertz >= 45 And CalibHertz <= 100000 Then
                                            CanDoIt = 1
                                        Else
                                            CalibratorSpecs "ACV", CalibParam, "V", CalibHertz / 1000, "kHz"
                                        End If
                                        
                                    Case "kHz"
                                        If CalibHertz >= 0.045 And CalibHertz <= 100 Then
                                            CanDoIt = 1
                                        Else
                                            MsgBox "Calibrator Is unable To Source " & CalibParam & CalibParamUnit & " at " & CalibHertz & CalibHertzUnit
                                        End If
                                        
                                End Select
                                CanDoIt = 1
                                
                            ElseIf CalibParam >= 33 And CalibParam <= 329.9999 Then
                                Select Case CalibHertzUnit
                                    Case "Hz"
                                        If CalibHertz >= 45 And CalibHertz <= 20000 Then
                                            CanDoIt = 1
                                        Else
                                            CalibratorSpecs "ACV", CalibParam, "V", CalibHertz / 1000, "kHz"
                                        End If
                                        
                                    Case "kHz"
                                        If CalibHertz >= 0.045 And CalibHertz <= 20 Then
                                            CanDoIt = 1
                                        Else
                                            MsgBox "Calibrator Is unable To Source " & CalibParam & CalibParamUnit & " at " & CalibHertz & CalibHertzUnit
                                        End If
                                        
                                End Select
                                CanDoIt = 1
                                
                            ElseIf CalibParam >= 330 And CalibParam <= 1020 Then
                                Select Case CalibHertzUnit
                                    Case "Hz"
                                        If CalibHertz >= 45 And CalibHertz <= 10000 Then
                                            CanDoIt = 1
                                        Else
                                            CalibratorSpecs "ACV", CalibParam, "V", CalibHertz / 1000, "kHz"
                                        End If
                                        
                                    Case "kHz"
                                        If CalibHertz >= 0.045 And CalibHertz <= 10 Then
                                            CanDoIt = 1
                                        Else
                                            MsgBox "Calibrator Is unable To Source " & CalibParam & CalibParamUnit & " at " & CalibHertz & CalibHertzUnit
                                        End If
                                        
                                End Select
                                CanDoIt = 1
                            
                            
                            Else
                                
                                CalibratorSpecs "ACV", CalibParam / 1000, "kV", CalibHertz, CalibHertzUnit
                            End If

                    Case "kV"
                            
                            If CalibParam >= 0.000001 And CalibParam <= 1.02 Then
                                CanDoIt = 1
                                
                            Else
                                MsgBox "Calibrator Is unable To Source " & CalibParam & CalibParamUnit & " at " & CalibHertz & CalibHertzUnit
                                CanDoIt = 0
                            End If
                            
                    End Select
'------------------------End ACV Check--------------------------------
                    
                    '------------------------Begin DCA Check------------------------------
                Case "DCI"
                    
                    Select Case CalibParamUnit
                        Case "uA"
                            If CalibParam >= -999.9999 And CalibParam <= 999.9999 Then
                                CanDoIt = 1
                                
                            Else
                                
                                CalibratorSpecs "DCA", CalibParam / 1000, "mA", "", ""
                            End If
                            
                        Case "mA"
                            If CalibParam >= -999.9999 And CalibParam <= 999.9999 Then
                                CanDoIt = 1
                                
                            Else
                                
                                CalibratorSpecs "DCA", CalibParam / 1000, "A", "", ""
                            End If
                        Case "A"
                            
                            If CalibParam >= -11 And CalibParam <= 11 Then
                                CanDoIt = 1
                                
                            Else
                                MsgBox "Calibrator Is unable To Source " & CalibParam & CalibParamUnit
                                CanDoIt = 0
                            End If
                            
                        Case "kA"
                            
                            MsgBox "Calibrator Is unable To Source " & CalibParam & CalibParamUnit
                            
                            '------------------------End DCA Check--------------------------------
                    End Select
                    
                    '------------------------Begin Ohm Check------------------------------
                Case "Ohm"
                    
                    Select Case CalibParamUnit
                        Case "uOhm"
                            ' MsgBox "uOhm"
                            If CalibParam >= 0 And CalibParam <= 999.9999 Then
                                CanDoIt = 1
                                
                            Else
                                
                                CalibratorSpecs "Ohm", CalibParam / 1000, "mOhm", "", ""
                            End If
                            
                        Case "mOhm"
                         'MsgBox "mOhm"
                            If CalibParam >= 0 And CalibParam <= 999.9999 Then
                                CanDoIt = 1
                                
                            Else
                                
                                CalibratorSpecs "Ohm", CalibParam / 1000, "Ohm", "", ""
                            End If
                            
                        Case "Ohm"
                            ' MsgBox "Ohm"
                            If CalibParam >= 0 And CalibParam <= 999.9999 Then
                                CanDoIt = 1
                                
                            Else
                                CalibratorSpecs "Ohm", CalibParam / 1000, "kOhm", "", ""
                            End If
                            
                        Case "kOhm"
                             'MsgBox "kOhm"
                            If CalibParam >= 0 And CalibParam <= 999.9999 Then
                                CanDoIt = 1
                               
                            Else
                                CalibratorSpecs "Ohm", CalibParam / 1000, "MOhm", "", ""
                                CanDoIt = 0
                            End If
                            
                        Case "MOhm"
                            'MsgBox "MOhm"
                            If CalibParam >= 0 And CalibParam <= 330 Then
                                CanDoIt = 1
                                
                            Else
                                MsgBox "Calibrator Is unable To Source " & CalibParam & CalibParamUnit
                                CanDoIt = 0
                            End If
                            
                        '------------------------End Ohm Check--------------------------------
                    
                Case "Temp"
'------------------------Begin Temp Check in C Degs------------------------------
                    
                    Select Case CalibParamUnit
                        Case "B"
                            If CalibParam >= 600 And CalibParam <= 1820 Then
                                CanDoIt = 1
                            Else
                                MsgBox "Calibrator Is unable To Source " & CalibParam & CalibParamUnit
                                CanDoIt = 0
                            End If
                            
                        Case "C"
                            If CalibParam >= 0 And CalibParam <= 2316 Then
                                CanDoIt = 1
                            Else
                                MsgBox "Calibrator Is unable To Source " & CalibParam & CalibParamUnit
                                CanDoIt = 0
                            End If
                        Case "E"
                            If CalibParam >= -250 And CalibParam <= 1000 Then
                                CanDoIt = 1
                            Else
                                MsgBox "Calibrator Is unable To Source " & CalibParam & CalibParamUnit
                                CanDoIt = 0
                            End If
                            
                        Case "J"
                            If CalibParam >= -210 And CalibParam <= 1200 Then
                                CanDoIt = 1
                            Else
                                MsgBox "Calibrator Is unable To Source " & CalibParam & CalibParamUnit
                                CanDoIt = 0
                            End If
                            
                        Case "K"
                            If CalibParam >= -200 And CalibParam <= 1372 Then
                                CanDoIt = 1
                            Else
                                MsgBox "Calibrator Is unable To Source " & CalibParam & CalibParamUnit
                                CanDoIt = 0
                            End If
                            
                        Case "L"
                            If CalibParam >= -200 And CalibParam <= 900 Then
                                CanDoIt = 1
                            Else
                                MsgBox "Calibrator Is unable To Source " & CalibParam & CalibParamUnit
                                CanDoIt = 0
                            End If
                            
                        Case "N"
                            If CalibParam >= -200 And CalibParam <= 1300 Then
                                CanDoIt = 1
                            Else
                                MsgBox "Calibrator Is unable To Source " & CalibParam & CalibParamUnit
                                CanDoIt = 0
                            End If
                            
                        Case "R"
                            If CalibParam >= 0 And CalibParam <= 1767 Then
                                CanDoIt = 1
                            Else
                                MsgBox "Calibrator Is unable To Source " & CalibParam & CalibParamUnit
                                CanDoIt = 0
                            End If
                            
                        Case "S"
                            If CalibParam >= 0 And CalibParam <= 1767 Then
                                CanDoIt = 1
                            Else
                                MsgBox "Calibrator Is unable To Source " & CalibParam & CalibParamUnit
                                CanDoIt = 0
                            End If
                            
                        Case "T"
                            If CalibParam >= -250 And CalibParam <= 400 Then
                                CanDoIt = 1
                            Else
                                MsgBox "Calibrator Is unable To Source " & CalibParam & CalibParamUnit
                                CanDoIt = 0
                            End If
                            
                        Case "U"
                            If CalibParam >= -200 And CalibParam <= 600 Then
                                CanDoIt = 1
                            Else
                                MsgBox "Calibrator Is unable To Source " & CalibParam & CalibParamUnit
                                CanDoIt = 0
                            End If
                            
                    End Select
'------------------------End Temp Check in C Degs--------------------------------

                
                Case "Cap"
'------------------------Begin Cap Check------------------------------
                    
                    Select Case CalibParamUnit
                        Case "pF"
                            If CalibParam >= 330 And CalibParam <= 999.9999 Then
                                CanDoIt = 1
                            Else
                                
                                CalibratorSpecs "Cap", CalibParam / 1000, "nF", "", ""
                            End If
                            
                        Case "nF"
                            If CalibParam >= 1 And CalibParam <= 999.9999 Then
                                CanDoIt = 1
                            Else
                                
                                CalibratorSpecs "Cap", CalibParam / 1000, "uF", "", ""
                            End If
                        Case "uF"
                            
                            If CalibParam >= 1 And CalibParam <= 999.999 Then
                                CanDoIt = 1
                            Else
                                CalibratorSpecs "Cap", CalibParam / 1000, "mF", "", ""
                                
                            End If
                            
                        Case "mF"
                            
                            If CalibParam >= 1 And CalibParam <= 1.1 Then
                                CanDoIt = 1
                                
                            Else
                                MsgBox "Calibrator Is unable To Source " & CalibParam & CalibParamUnit
                                CanDoIt = 0
                            End If
                            
                    End Select
'------------------------End Cap Check--------------------------------

                        
                    End Select
                    
            End Select
            
        Case "5502A"
            Select Case CalFunc
                Case "DCV"
                    '------------------------Begin DCV Check------------------------------
                    
                    Select Case CalibParamUnit
                        Case "uV"
                            If CalibParam >= -999.9999 And CalibParam <= 999.9999 Then
                                CanDoIt = 1
                            Else
                                
                                CalibratorSpecs "DCV", CalibParam / 1000, "mV", "", ""
                            End If
                            
                        Case "mV"
                            If CalibParam >= -999.9999 And CalibParam <= 999.9999 Then
                                CanDoIt = 1
                            Else
                                
                                CalibratorSpecs "DCV", CalibParam / 1000, "V", "", ""
                            End If
                        Case "V"
                            
                            If CalibParam >= -1000 And CalibParam <= 1000 Then
                                CanDoIt = 1
                            Else
                                CalibratorSpecs "DCV", CalibParam / 1000, "kV", "", ""
                            End If
                            
                        Case "kV"
                            
                            If CalibParam >= -1 And CalibParam <= 1 Then
                                CanDoIt = 1
                                
                            Else
                                MsgBox "Calibrator Is unable To Source " & CalibParam & CalibParamUnit
                                CanDoIt = 0
                            End If
                            '------------------------End DCV Check--------------------------------
                    End Select
                    
                    
                Case "ACV"
'------------------------Begin ACV Check------------------------------
                    
                    Select Case CalibParamUnit
                        Case "mV"
                            If CalibParam >= 1 And CalibParam <= 329.999 Then         '-----------------Check to see if Calibrator can go lower voltage than 1 .1,.001
                                Select Case CalibHertzUnit
                                    Case "Hz"
                                        If CalibHertz >= 10 And CalibHertz <= 500000 Then
                                            CanDoIt = 1
                                        Else
                                            CalibratorSpecs "ACV", CalibParam, "mV", CalibHertz / 1000, "kHz"
                                        End If
                                        
                                    Case "kHz"
                                        If CalibHertz >= 0.01 And CalibHertz <= 500 Then
                                            CanDoIt = 1
                                        Else
                                            MsgBox "Calibrator Is unable To Source " & CalibParam & CalibParamUnit & " at " & CalibHertz & CalibHertzUnit
                                        End If
                                        
                                End Select
                                CanDoIt = 1
                            Else
                                
                                CalibratorSpecs "ACV", CalibParam / 1000, "V", CalibHertz, CalibHertzUnit
                            End If
                            
                        Case "V"
                            If CalibParam >= 0.001 And CalibParam <= 3.2999 Then
                                Select Case CalibHertzUnit
                                    Case "Hz"
                                        If CalibHertz >= 10 And CalibHertz <= 500000 Then
                                            CanDoIt = 1
                                        Else
                                            CalibratorSpecs "ACV", CalibParam, "V", CalibHertz / 1000, "kHz"
                                        End If
                                        
                                    Case "kHz"
                                        If CalibHertz >= 0.01 And CalibHertz <= 500 Then
                                            CanDoIt = 1
                                        Else
                                            MsgBox "Calibrator Is unable To Source " & CalibParam & CalibParamUnit & " at " & CalibHertz & CalibHertzUnit
                                        End If
                                        
                                End Select
                                CanDoIt = 1
                                
                            ElseIf CalibParam >= 3.3 And CalibParam <= 32.9999 Then
                                Select Case CalibHertzUnit
                                    Case "Hz"
                                        If CalibHertz >= 10 And CalibHertz <= 100000 Then
                                            CanDoIt = 1
                                        Else
                                            CalibratorSpecs "ACV", CalibParam, "V", CalibHertz / 1000, "kHz"
                                        End If
                                        
                                    Case "kHz"
                                        If CalibHertz >= 0.01 And CalibHertz <= 100 Then
                                            CanDoIt = 1
                                        Else
                                            MsgBox "Calibrator Is unable To Source " & CalibParam & CalibParamUnit & " at " & CalibHertz & CalibHertzUnit
                                        End If
                                        
                                End Select
                                CanDoIt = 1
                                
                            ElseIf CalibParam >= 33 And CalibParam <= 329.9999 Then
                                Select Case CalibHertzUnit
                                    Case "Hz"
                                        If CalibHertz >= 45 And CalibHertz <= 100000 Then
                                            CanDoIt = 1
                                        Else
                                            CalibratorSpecs "ACV", CalibParam, "V", CalibHertz / 1000, "kHz"
                                        End If
                                        
                                    Case "kHz"
                                        If CalibHertz >= 0.045 And CalibHertz <= 100 Then
                                            CanDoIt = 1
                                        Else
                                            MsgBox "Calibrator Is unable To Source " & CalibParam & CalibParamUnit & " at " & CalibHertz & CalibHertzUnit
                                        End If
                                        
                                End Select
                                CanDoIt = 1
                                
                            ElseIf CalibParam >= 330 And CalibParam <= 1020 Then
                                Select Case CalibHertzUnit
                                    Case "Hz"
                                        If CalibHertz >= 45 And CalibHertz <= 10000 Then
                                            CanDoIt = 1
                                        Else
                                            CalibratorSpecs "ACV", CalibParam, "V", CalibHertz / 1000, "kHz"
                                        End If
                                        
                                    Case "kHz"
                                        If CalibHertz >= 0.045 And CalibHertz <= 10 Then
                                            CanDoIt = 1
                                        Else
                                            MsgBox "Calibrator Is unable To Source " & CalibParam & CalibParamUnit & " at " & CalibHertz & CalibHertzUnit
                                        End If
                                        
                                End Select
                                CanDoIt = 1
                            
                            
                            Else
                                
                                CalibratorSpecs "ACV", CalibParam / 1000, "kV", CalibHertz, CalibHertzUnit
                            End If

                    Case "kV"
                            
                            If CalibParam >= 0.000001 And CalibParam <= 1020 Then
                                CanDoIt = 1
                                
                            Else
                                MsgBox "Calibrator Is unable To Source " & CalibParam & CalibParamUnit & " at " & CalibHertz & CalibHertzUnit
                                CanDoIt = 0
                            End If
                            
                    End Select
'------------------------End ACV Check--------------------------------
                    
                    '------------------------Begin DCI Check------------------------------
                Case "DCI"
                    
                    Select Case CalibParamUnit
                        Case "uA"
                            If CalibParam >= -999.9999 And CalibParam <= 999.9999 Then
                                CanDoIt = 1
                                
                            Else
                                
                                CalibratorSpecs "DCA", CalibParam / 1000, "mA", "", ""
                            End If
                            
                        Case "mA"
                            If CalibParam >= -999.9999 And CalibParam <= 999.9999 Then
                                CanDoIt = 1
                                
                            Else
                                
                                CalibratorSpecs "DCA", CalibParam / 1000, "A", "", ""
                            End If
                        Case "A"
                            
                            If CalibParam >= -20.5 And CalibParam <= 20.5 Then
                                CanDoIt = 1
                                
                            Else
                                MsgBox "Calibrator Is unable To Source " & CalibParam & CalibParamUnit
                                CanDoIt = 0
                            End If
                            
                        Case "kA"
                            
                            MsgBox "Calibrator Is unable To Source " & CalibParam & CalibParamUnit
                            
                            '------------------------End DCI Check--------------------------------
                    End Select
                    
'------------------------Begin ACI Check------------------------------
                Case "ACI"
                    MsgBox "Here"
                    Select Case CalibParamUnit
                        Case "uA"
                            If CalibParam >= 1 And CalibParam <= 329.999 Then         '-----------------Check to see if Calibrator can go lower voltage than 1 .1,.001
                                Select Case CalibHertzUnit
                                    Case "Hz"
                                        If CalibHertz >= 10 And CalibHertz <= 30000 Then
                                            CanDoIt = 1
                                        Else
                                            CalibratorSpecs "ACI", CalibParam, "uA", CalibHertz / 1000, "kHz"
                                        End If
                                        
                                    Case "kHz"
                                        If CalibHertz >= 0.01 And CalibHertz <= 30 Then
                                            CanDoIt = 1
                                        Else
                                            MsgBox "Calibrator Is unable To Source " & CalibParam & CalibParamUnit & " at " & CalibHertz & CalibHertzUnit
                                        End If
                                        
                                End Select
                                CanDoIt = 1
                            Else
                                
                                CalibratorSpecs "ACI", CalibParam / 1000, "mA", CalibHertz, CalibHertzUnit
                            End If
                            
                        Case "mA"
                            If CalibParam >= 0.001 And CalibParam <= 3.2999 Then
                                Select Case CalibHertzUnit
                                    Case "Hz"
                                        If CalibHertz >= 10 And CalibHertz <= 30000 Then
                                            CanDoIt = 1
                                        Else
                                            CalibratorSpecs "ACI", CalibParam, "mA", CalibHertz / 1000, "kHz"
                                        End If
                                        
                                    Case "kHz"
                                        If CalibHertz >= 0.01 And CalibHertz <= 30 Then
                                            CanDoIt = 1
                                        Else
                                            MsgBox "Calibrator Is unable To Source " & CalibParam & CalibParamUnit & " at " & CalibHertz & CalibHertzUnit
                                        End If
                                        
                                End Select
                                CanDoIt = 1
                                
                            ElseIf CalibParam >= 3.3 And CalibParam <= 32.9999 Then
                                Select Case CalibHertzUnit
                                    Case "Hz"
                                        If CalibHertz >= 10 And CalibHertz <= 30000 Then
                                            CanDoIt = 1
                                        Else
                                            CalibratorSpecs "ACI", CalibParam, "mA", CalibHertz / 1000, "kHz"
                                        End If
                                        
                                    Case "kHz"
                                        If CalibHertz >= 0.01 And CalibHertz <= 30 Then
                                            CanDoIt = 1
                                        Else
                                            MsgBox "Calibrator Is unable To Source " & CalibParam & CalibParamUnit & " at " & CalibHertz & CalibHertzUnit
                                        End If
                                        
                                End Select
                                CanDoIt = 1
                                
                            ElseIf CalibParam >= 33 And CalibParam <= 329.9999 Then
                                Select Case CalibHertzUnit
                                    Case "Hz"
                                        If CalibHertz >= 10 And CalibHertz <= 30000 Then
                                            CanDoIt = 1
                                        Else
                                            CalibratorSpecs "ACI", CalibParam, "mA", CalibHertz / 1000, "kHz"
                                        End If
                                        
                                    Case "kHz"
                                        If CalibHertz >= 0.01 And CalibHertz <= 30 Then
                                            CanDoIt = 1
                                        Else
                                            MsgBox "Calibrator Is unable To Source " & CalibParam & CalibParamUnit & " at " & CalibHertz & CalibHertzUnit
                                        End If
                                        
                                End Select
                                CanDoIt = 1
                                
                            ElseIf CalibParam >= 330 And CalibParam <= 1099.99 Then
                                Select Case CalibHertzUnit
                                    Case "Hz"
                                        If CalibHertz >= 10 And CalibHertz <= 10000 Then
                                            CanDoIt = 1
                                        Else
                                            CalibratorSpecs "ACI", CalibParam, "mA", CalibHertz / 1000, "kHz"
                                        End If
                                        
                                    Case "kHz"
                                        If CalibHertz >= 0.01 And CalibHertz <= 10 Then
                                            CanDoIt = 1
                                        Else
                                            MsgBox "Calibrator Is unable To Source " & CalibParam & CalibParamUnit & " at " & CalibHertz & CalibHertzUnit
                                        End If
                                        
                                End Select
                                CanDoIt = 1
                            
                            
                            Else
                                
                                CalibratorSpecs "ACI", CalibParam / 1000, "kV", CalibHertz, CalibHertzUnit
                            End If

                    Case "A"
                            
                            If CalibParam >= 1 And CalibParam <= 2.99999 Then
                                
                                Select Case CalibHertzUnit
                                    Case "Hz"
                                        If CalibHertz >= 10 And CalibHertz <= 10000 Then
                                            CanDoIt = 1
                                        Else
                                            CalibratorSpecs "ACI", CalibParam, "A", CalibHertz / 1000, "kHz"
                                        End If
                                        
                                    Case "kHz"
                                        If CalibHertz >= 0.01 And CalibHertz <= 10 Then
                                            CanDoIt = 1
                                        Else
                                            MsgBox "Calibrator Is unable To Source " & CalibParam & CalibParamUnit & " at " & CalibHertz & CalibHertzUnit
                                        End If
                                        
                                End Select
                                CanDoIt = 1
                                
                            ElseIf CalibParam >= 3 And CalibParam <= 10.9999 Then
                                Select Case CalibHertzUnit
                                    Case "Hz"
                                        If CalibHertz >= 45 And CalibHertz <= 5000 Then
                                            CanDoIt = 1
                                        Else
                                            CalibratorSpecs "ACI", CalibParam, "A", CalibHertz / 1000, "kHz"
                                        End If
                                        
                                    Case "kHz"
                                        If CalibHertz >= 0.045 And CalibHertz <= 5 Then
                                            CanDoIt = 1
                                        Else
                                            MsgBox "Calibrator Is unable To Source " & CalibParam & CalibParamUnit & " at " & CalibHertz & CalibHertzUnit
                                        End If
                                        
                                End Select
                                CanDoIt = 1
                                
                            ElseIf CalibParam >= 11 And CalibParam <= 20.5 Then
                                Select Case CalibHertzUnit
                                    Case "Hz"
                                        If CalibHertz >= 45 And CalibHertz <= 5000 Then
                                            CanDoIt = 1
                                        Else
                                            CalibratorSpecs "ACI", CalibParam, "A", CalibHertz / 1000, "kHz"
                                        End If
                                        
                                    Case "kHz"
                                        If CalibHertz >= 0.045 And CalibHertz <= 5 Then
                                            CanDoIt = 1
                                        Else
                                            MsgBox "Calibrator Is unable To Source " & CalibParam & CalibParamUnit & " at " & CalibHertz & CalibHertzUnit
                                        End If
                                        
                                End Select
                                CanDoIt = 1
                                
                            Else
                                MsgBox "Calibrator Is unable To Source " & CalibParam & CalibParamUnit & " at " & CalibHertz & CalibHertzUnit
                                CanDoIt = 0
                            End If
                            
                    End Select
'------------------------End ACI Check--------------------------------

                    
                    '------------------------Begin Ohm Check------------------------------
                Case "Ohm"
                    
                    Select Case CalibParamUnit
                        Case "uOhm"
                             'MsgBox "uOhm"
                            If CalibParam >= 0 And CalibParam <= 999.9999 Then
                                CanDoIt = 1
                                
                            Else
                                
                                CalibratorSpecs "Ohm", CalibParam / 1000, "mOhm", "", ""
                            End If
                            
                        Case "mOhm"
                        ' MsgBox "mOhm"
                            If CalibParam >= 0 And CalibParam <= 999.9999 Then
                                CanDoIt = 1
                                
                            Else
                                
                                CalibratorSpecs "Ohm", CalibParam / 1000, "Ohm", "", ""
                            End If
                            
                        Case "Ohm"
                             'MsgBox "Ohm"
                            If CalibParam >= 0 And CalibParam <= 999.9999 Then
                                CanDoIt = 1
                                
                            Else
                                CalibratorSpecs "Ohm", CalibParam / 1000, "kOhm", "", ""
                            End If
                            
                        Case "kOhm"
                             'MsgBox "kOhm"
                            If CalibParam >= 0 And CalibParam <= 999.9999 Then
                                CanDoIt = 1
                               
                            Else
                                CalibratorSpecs "Ohm", CalibParam / 1000, "MOhm", "", ""
                                CanDoIt = 0
                            End If
                            
                        Case "MOhm"
                           ' MsgBox "MOhm"
                            If CalibParam >= 0 And CalibParam <= 999.9999 Then
                                CanDoIt = 1
                                
                            Else
                                CalibratorSpecs "Ohm", CalibParam / 1000, "GOhm", "", ""
                                CanDoIt = 0
                            End If
                            
                        Case "GOhm"
                            
                            If CalibParam >= 0 And CalibParam <= 1.1 Then
                                CanDoIt = 1
                                
                            Else
                                MsgBox "Calibrator Is unable To Source " & CalibParam & CalibParamUnit
                                CanDoIt = 0
                            End If
                            
                            '------------------------End Ohm Check--------------------------------
                    End Select
                    
                Case "Temp"
'------------------------Begin Temp Check in C Degs------------------------------
                    
                    Select Case CalibParamUnit
                        Case "B"
                            If CalibParam >= 600 And CalibParam <= 1820 Then
                                CanDoIt = 1
                            Else
                                MsgBox "Calibrator Is unable To Source " & CalibParam & CalibParamUnit
                                CanDoIt = 0
                            End If
                            
                        Case "C"
                            If CalibParam >= 0 And CalibParam <= 2316 Then
                                CanDoIt = 1
                            Else
                                MsgBox "Calibrator Is unable To Source " & CalibParam & CalibParamUnit
                                CanDoIt = 0
                            End If
                        Case "E"
                            If CalibParam >= -250 And CalibParam <= 1000 Then
                                CanDoIt = 1
                            Else
                                MsgBox "Calibrator Is unable To Source " & CalibParam & CalibParamUnit
                                CanDoIt = 0
                            End If
                            
                        Case "J"
                            If CalibParam >= -210 And CalibParam <= 1200 Then
                                CanDoIt = 1
                            Else
                                MsgBox "Calibrator Is unable To Source " & CalibParam & CalibParamUnit
                                CanDoIt = 0
                            End If
                            
                        Case "K"
                            If CalibParam >= -200 And CalibParam <= 1372 Then
                                CanDoIt = 1
                            Else
                                MsgBox "Calibrator Is unable To Source " & CalibParam & CalibParamUnit
                                CanDoIt = 0
                            End If
                            
                        Case "L"
                            If CalibParam >= -200 And CalibParam <= 900 Then
                                CanDoIt = 1
                            Else
                                MsgBox "Calibrator Is unable To Source " & CalibParam & CalibParamUnit
                                CanDoIt = 0
                            End If
                            
                        Case "N"
                            If CalibParam >= -200 And CalibParam <= 1300 Then
                                CanDoIt = 1
                            Else
                                MsgBox "Calibrator Is unable To Source " & CalibParam & CalibParamUnit
                                CanDoIt = 0
                            End If
                            
                        Case "R"
                            If CalibParam >= 0 And CalibParam <= 1767 Then
                                CanDoIt = 1
                            Else
                                MsgBox "Calibrator Is unable To Source " & CalibParam & CalibParamUnit
                                CanDoIt = 0
                            End If
                            
                        Case "S"
                            If CalibParam >= 0 And CalibParam <= 1767 Then
                                CanDoIt = 1
                            Else
                                MsgBox "Calibrator Is unable To Source " & CalibParam & CalibParamUnit
                                CanDoIt = 0
                            End If
                            
                        Case "T"
                            If CalibParam >= -250 And CalibParam <= 400 Then
                                CanDoIt = 1
                            Else
                                MsgBox "Calibrator Is unable To Source " & CalibParam & CalibParamUnit
                                CanDoIt = 0
                            End If
                            
                        Case "U"
                            If CalibParam >= -200 And CalibParam <= 600 Then
                                CanDoIt = 1
                            Else
                                MsgBox "Calibrator Is unable To Source " & CalibParam & CalibParamUnit
                                CanDoIt = 0
                            End If
                            
                    End Select
'------------------------End Temp Check in C Degs--------------------------------
                
                Case "Cap"
'------------------------Begin Cap Check------------------------------
                    
                    Select Case CalibParamUnit
                        Case "pF"
                            If CalibParam >= 220 And CalibParam <= 999.9999 Then
                                CanDoIt = 1
                            Else
                                
                                CalibratorSpecs "Cap", CalibParam / 1000, "nF", "", ""
                            End If
                            
                        Case "nF"
                            If CalibParam >= 1 And CalibParam <= 999.9999 Then
                                CanDoIt = 1
                            Else
                                
                                CalibratorSpecs "Cap", CalibParam / 1000, "uF", "", ""
                            End If
                        Case "uF"
                            
                            If CalibParam >= 1 And CalibParam <= 999.999 Then
                                CanDoIt = 1
                            Else
                                CalibratorSpecs "Cap", CalibParam / 1000, "mF", "", ""
                                
                            End If
                            
                        Case "mF"
                            
                            If CalibParam >= 1 And CalibParam <= 110 Then
                                CanDoIt = 1
                                
                            Else
                                MsgBox "Calibrator Is unable To Source " & CalibParam & CalibParamUnit
                                CanDoIt = 0
                            End If
                            
                    End Select
'------------------------End Cap Check--------------------------------
                    
            End Select
            
        Case "5522A"
            Select Case CalFunc
                Case "DCV"
                    '------------------------Begin DCV Check------------------------------
                    
                    Select Case CalibParamUnit
                        Case "uV"
                            If CalibParam >= -999.9999 And CalibParam <= 999.9999 Then
                                CanDoIt = 1
                            Else
                                
                                CalibratorSpecs "DCV", CalibParam / 1000, "mV", "", ""
                            End If
                            
                        Case "mV"
                            If CalibParam >= -999.9999 And CalibParam <= 999.9999 Then
                                CanDoIt = 1
                            Else
                                
                                CalibratorSpecs "DCV", CalibParam / 1000, "V", "", ""
                            End If
                        Case "V"
                            
                            If CalibParam >= -1020 And CalibParam <= 1020 Then
                                CanDoIt = 1
                            Else
                                CalibratorSpecs "DCV", CalibParam / 1000, "kV", "", ""
                            End If
                            
                        Case "kV"
                            
                            If CalibParam >= -1 And CalibParam <= 1 Then
                                CanDoIt = 1
                                
                            Else
                                MsgBox "Calibrator Is unable To Source " & CalibParam & CalibParamUnit
                                CanDoIt = 0
                            End If
                            '------------------------End DCV Check--------------------------------
                    End Select
                    
                    
                Case "ACV"
'------------------------Begin ACV Check------------------------------
                    
                    Select Case CalibParamUnit
                        Case "mV"
                            If CalibParam >= 1 And CalibParam <= 329.999 Then         '-----------------Check to see if Calibrator can go lower voltage than 1 .1,.001
                                Select Case CalibHertzUnit
                                    Case "Hz"
                                        If CalibHertz >= 10 And CalibHertz <= 500000 Then
                                            CanDoIt = 1
                                        Else
                                            CalibratorSpecs "ACV", CalibParam, "mV", CalibHertz / 1000, "kHz"
                                        End If
                                        
                                    Case "kHz"
                                        If CalibHertz >= 0.01 And CalibHertz <= 500 Then
                                            CanDoIt = 1
                                        Else
                                            MsgBox "Calibrator Is unable To Source " & CalibParam & CalibParamUnit & " at " & CalibHertz & CalibHertzUnit
                                        End If
                                        
                                End Select
                                CanDoIt = 1
                            Else
                                
                                CalibratorSpecs "ACV", CalibParam / 1000, "V", CalibHertz, CalibHertzUnit
                            End If
                            
                        Case "V"
                            If CalibParam >= 0.001 And CalibParam <= 3.2999 Then
                                Select Case CalibHertzUnit
                                    Case "Hz"
                                        If CalibHertz >= 10 And CalibHertz <= 500000 Then
                                            CanDoIt = 1
                                        Else
                                            CalibratorSpecs "ACV", CalibParam, "V", CalibHertz / 1000, "kHz"
                                        End If
                                        
                                    Case "kHz"
                                        If CalibHertz >= 0.01 And CalibHertz <= 500 Then
                                            CanDoIt = 1
                                        Else
                                            MsgBox "Calibrator Is unable To Source " & CalibParam & CalibParamUnit & " at " & CalibHertz & CalibHertzUnit
                                        End If
                                        
                                End Select
                                CanDoIt = 1
                                
                            ElseIf CalibParam >= 3.3 And CalibParam <= 32.9999 Then
                                Select Case CalibHertzUnit
                                    Case "Hz"
                                        If CalibHertz >= 10 And CalibHertz <= 100000 Then
                                            CanDoIt = 1
                                        Else
                                            CalibratorSpecs "ACV", CalibParam, "V", CalibHertz / 1000, "kHz"
                                        End If
                                        
                                    Case "kHz"
                                        If CalibHertz >= 0.01 And CalibHertz <= 100 Then
                                            CanDoIt = 1
                                        Else
                                            MsgBox "Calibrator Is unable To Source " & CalibParam & CalibParamUnit & " at " & CalibHertz & CalibHertzUnit
                                        End If
                                        
                                End Select
                                CanDoIt = 1
                                
                            ElseIf CalibParam >= 33 And CalibParam <= 329.9999 Then
                                Select Case CalibHertzUnit
                                    Case "Hz"
                                        If CalibHertz >= 45 And CalibHertz <= 100000 Then
                                            CanDoIt = 1
                                        Else
                                            CalibratorSpecs "ACV", CalibParam, "V", CalibHertz / 1000, "kHz"
                                        End If
                                        
                                    Case "kHz"
                                        If CalibHertz >= 0.045 And CalibHertz <= 100 Then
                                            CanDoIt = 1
                                        Else
                                            MsgBox "Calibrator Is unable To Source " & CalibParam & CalibParamUnit & " at " & CalibHertz & CalibHertzUnit
                                        End If
                                        
                                End Select
                                CanDoIt = 1
                                
                            ElseIf CalibParam >= 330 And CalibParam <= 1020 Then
                                Select Case CalibHertzUnit
                                    Case "Hz"
                                        If CalibHertz >= 45 And CalibHertz <= 10000 Then
                                            CanDoIt = 1
                                        Else
                                            CalibratorSpecs "ACV", CalibParam, "V", CalibHertz / 1000, "kHz"
                                        End If
                                        
                                    Case "kHz"
                                        If CalibHertz >= 0.045 And CalibHertz <= 10 Then
                                            CanDoIt = 1
                                        Else
                                            MsgBox "Calibrator Is unable To Source " & CalibParam & CalibParamUnit & " at " & CalibHertz & CalibHertzUnit
                                        End If
                                        
                                End Select
                                CanDoIt = 1
                            
                            
                            Else
                                
                                CalibratorSpecs "ACV", CalibParam / 1000, "kV", CalibHertz, CalibHertzUnit
                            End If

                    Case "kV"
                            
                            If CalibParam >= 0.000001 And CalibParam <= 1.02 Then
                                CanDoIt = 1
                                
                            Else
                                MsgBox "Calibrator Is unable To Source " & CalibParam & CalibParamUnit & " at " & CalibHertz & CalibHertzUnit
                                CanDoIt = 0
                            End If
                            
                    End Select
'------------------------End ACV Check--------------------------------
                    
'------------------------Begin DCI Check------------------------------
                Case "DCI"
                    
                    Select Case CalibParamUnit
                        Case "uA"
                            If CalibParam >= -999.9999 And CalibParam <= 999.9999 Then
                                CanDoIt = 1
                                
                            Else
                                
                                CalibratorSpecs "DCI", CalibParam / 1000, "mA", "", ""
                            End If
                            
                        Case "mA"
                            If CalibParam >= -999.9999 And CalibParam <= 999.9999 Then
                                CanDoIt = 1
                                
                            Else
                                
                                CalibratorSpecs "DCI", CalibParam / 1000, "A", "", ""
                            End If
                        Case "A"
                            
                            If CalibParam >= -20.5 And CalibParam <= 20.5 Then
                                CanDoIt = 1
                                
                            Else
                                MsgBox "Calibrator Is unable To Source " & CalibParam & CalibParamUnit
                                CanDoIt = 0
                            End If
                            
                        Case "kA"
                            
                            MsgBox "Calibrator Is unable To Source " & CalibParam & CalibParamUnit
                            
'------------------------End DCA Check--------------------------------
                    End Select
                    
'------------------------Begin ACI Check------------------------------
                Case "ACI"
                    MsgBox "Here"
                    Select Case CalibParamUnit
                        Case "uA"
                            If CalibParam >= 1 And CalibParam <= 329.999 Then         '-----------------Check to see if Calibrator can go lower voltage than 1 .1,.001
                                Select Case CalibHertzUnit
                                    Case "Hz"
                                        If CalibHertz >= 10 And CalibHertz <= 30000 Then
                                            CanDoIt = 1
                                        Else
                                            CalibratorSpecs "ACI", CalibParam, "uA", CalibHertz / 1000, "kHz"
                                        End If
                                        
                                    Case "kHz"
                                        If CalibHertz >= 0.01 And CalibHertz <= 30 Then
                                            CanDoIt = 1
                                        Else
                                            MsgBox "Calibrator Is unable To Source " & CalibParam & CalibParamUnit & " at " & CalibHertz & CalibHertzUnit
                                        End If
                                        
                                End Select
                                CanDoIt = 1
                            Else
                                
                                CalibratorSpecs "ACI", CalibParam / 1000, "mA", CalibHertz, CalibHertzUnit
                            End If
                            
                        Case "mA"
                            If CalibParam >= 0.001 And CalibParam <= 3.2999 Then
                                Select Case CalibHertzUnit
                                    Case "Hz"
                                        If CalibHertz >= 10 And CalibHertz <= 30000 Then
                                            CanDoIt = 1
                                        Else
                                            CalibratorSpecs "ACI", CalibParam, "mA", CalibHertz / 1000, "kHz"
                                        End If
                                        
                                    Case "kHz"
                                        If CalibHertz >= 0.01 And CalibHertz <= 30 Then
                                            CanDoIt = 1
                                        Else
                                            MsgBox "Calibrator Is unable To Source " & CalibParam & CalibParamUnit & " at " & CalibHertz & CalibHertzUnit
                                        End If
                                        
                                End Select
                                CanDoIt = 1
                                
                            ElseIf CalibParam >= 3.3 And CalibParam <= 32.9999 Then
                                Select Case CalibHertzUnit
                                    Case "Hz"
                                        If CalibHertz >= 10 And CalibHertz <= 30000 Then
                                            CanDoIt = 1
                                        Else
                                            CalibratorSpecs "ACI", CalibParam, "mA", CalibHertz / 1000, "kHz"
                                        End If
                                        
                                    Case "kHz"
                                        If CalibHertz >= 0.01 And CalibHertz <= 30 Then
                                            CanDoIt = 1
                                        Else
                                            MsgBox "Calibrator Is unable To Source " & CalibParam & CalibParamUnit & " at " & CalibHertz & CalibHertzUnit
                                        End If
                                        
                                End Select
                                CanDoIt = 1
                                
                            ElseIf CalibParam >= 33 And CalibParam <= 329.9999 Then
                                Select Case CalibHertzUnit
                                    Case "Hz"
                                        If CalibHertz >= 10 And CalibHertz <= 30000 Then
                                            CanDoIt = 1
                                        Else
                                            CalibratorSpecs "ACI", CalibParam, "mA", CalibHertz / 1000, "kHz"
                                        End If
                                        
                                    Case "kHz"
                                        If CalibHertz >= 0.01 And CalibHertz <= 30 Then
                                            CanDoIt = 1
                                        Else
                                            MsgBox "Calibrator Is unable To Source " & CalibParam & CalibParamUnit & " at " & CalibHertz & CalibHertzUnit
                                        End If
                                        
                                End Select
                                CanDoIt = 1
                                
                            ElseIf CalibParam >= 330 And CalibParam <= 1099.99 Then
                                Select Case CalibHertzUnit
                                    Case "Hz"
                                        If CalibHertz >= 10 And CalibHertz <= 10000 Then
                                            CanDoIt = 1
                                        Else
                                            CalibratorSpecs "ACI", CalibParam, "mA", CalibHertz / 1000, "kHz"
                                        End If
                                        
                                    Case "kHz"
                                        If CalibHertz >= 0.01 And CalibHertz <= 10 Then
                                            CanDoIt = 1
                                        Else
                                            MsgBox "Calibrator Is unable To Source " & CalibParam & CalibParamUnit & " at " & CalibHertz & CalibHertzUnit
                                        End If
                                        
                                End Select
                                CanDoIt = 1
                            
                            
                            Else
                                
                                CalibratorSpecs "ACI", CalibParam / 1000, "kV", CalibHertz, CalibHertzUnit
                            End If

                    Case "A"
                            
                            If CalibParam >= 1 And CalibParam <= 2.99999 Then
                                
                                Select Case CalibHertzUnit
                                    Case "Hz"
                                        If CalibHertz >= 10 And CalibHertz <= 10000 Then
                                            CanDoIt = 1
                                        Else
                                            CalibratorSpecs "ACI", CalibParam, "A", CalibHertz / 1000, "kHz"
                                        End If
                                        
                                    Case "kHz"
                                        If CalibHertz >= 0.01 And CalibHertz <= 10 Then
                                            CanDoIt = 1
                                        Else
                                            MsgBox "Calibrator Is unable To Source " & CalibParam & CalibParamUnit & " at " & CalibHertz & CalibHertzUnit
                                        End If
                                        
                                End Select
                                CanDoIt = 1
                                
                            ElseIf CalibParam >= 3 And CalibParam <= 10.9999 Then
                                Select Case CalibHertzUnit
                                    Case "Hz"
                                        If CalibHertz >= 45 And CalibHertz <= 5000 Then
                                            CanDoIt = 1
                                        Else
                                            CalibratorSpecs "ACI", CalibParam, "A", CalibHertz / 1000, "kHz"
                                        End If
                                        
                                    Case "kHz"
                                        If CalibHertz >= 0.045 And CalibHertz <= 5 Then
                                            CanDoIt = 1
                                        Else
                                            MsgBox "Calibrator Is unable To Source " & CalibParam & CalibParamUnit & " at " & CalibHertz & CalibHertzUnit
                                        End If
                                        
                                End Select
                                CanDoIt = 1
                                
                            ElseIf CalibParam >= 11 And CalibParam <= 20.5 Then
                                Select Case CalibHertzUnit
                                    Case "Hz"
                                        If CalibHertz >= 45 And CalibHertz <= 5000 Then
                                            CanDoIt = 1
                                        Else
                                            CalibratorSpecs "ACI", CalibParam, "A", CalibHertz / 1000, "kHz"
                                        End If
                                        
                                    Case "kHz"
                                        If CalibHertz >= 0.045 And CalibHertz <= 5 Then
                                            CanDoIt = 1
                                        Else
                                            MsgBox "Calibrator Is unable To Source " & CalibParam & CalibParamUnit & " at " & CalibHertz & CalibHertzUnit
                                        End If
                                        
                                End Select
                                CanDoIt = 1
                                
                            Else
                                MsgBox "Calibrator Is unable To Source " & CalibParam & CalibParamUnit & " at " & CalibHertz & CalibHertzUnit
                                CanDoIt = 0
                            End If
                            
                    End Select
'------------------------End ACI Check--------------------------------

                    
'------------------------Begin Ohm Check------------------------------
                Case "Ohm"
                    
                    Select Case CalibParamUnit
                        Case "uOhm"
                            'MsgBox "uOhm"
                            If CalibParam >= 0 And CalibParam <= 999.9999 Then
                                CanDoIt = 1
                                
                            Else
                                
                                CalibratorSpecs "Ohm", CalibParam / 1000, "mOhm", "", ""
                            End If
                            
                        Case "mOhm"
                         'MsgBox "mOhm"
                            If CalibParam >= 0 And CalibParam <= 999.9999 Then
                                CanDoIt = 1
                                
                            Else
                                
                                CalibratorSpecs "Ohm", CalibParam / 1000, "Ohm", "", ""
                            End If
                            
                        Case "Ohm"
                             'MsgBox "Ohm"
                            If CalibParam >= 0 And CalibParam <= 999.9999 Then
                                CanDoIt = 1
                                
                            Else
                                CalibratorSpecs "Ohm", CalibParam / 1000, "kOhm", "", ""
                            End If
                            
                        Case "kOhm"
                             'MsgBox "kOhm"
                            If CalibParam >= 0 And CalibParam <= 999.9999 Then
                                CanDoIt = 1
                               
                            Else
                                CalibratorSpecs "Ohm", CalibParam / 1000, "MOhm", "", ""
                                CanDoIt = 0
                            End If
                            
                        Case "MOhm"
                            'MsgBox "MOhm"
                            If CalibParam >= 0 And CalibParam <= 999.9999 Then
                                CanDoIt = 1
                                
                            Else
                                CalibratorSpecs "Ohm", CalibParam / 1000, "GOhm", "", ""
                                CanDoIt = 0
                            End If
                            
                        Case "GOhm"
                            
                            If CalibParam >= 0 And CalibParam <= 1.1 Then
                                CanDoIt = 1
                                
                            Else
                                MsgBox "Calibrator Is unable To Source " & CalibParam & CalibParamUnit
                                CanDoIt = 0
                            End If
                            
                            '------------------------End Ohm Check--------------------------------
                    End Select
                    
                Case "Temp"
'------------------------Begin Temp Check in C Degs------------------------------
                    
                    Select Case CalibParamUnit
                        Case "B"
                            If CalibParam >= 600 And CalibParam <= 1820 Then
                                CanDoIt = 1
                            Else
                                MsgBox "Calibrator Is unable To Source " & CalibParam & CalibParamUnit
                                CanDoIt = 0
                            End If
                            
                        Case "C"
                            If CalibParam >= 0 And CalibParam <= 2316 Then
                                CanDoIt = 1
                            Else
                                MsgBox "Calibrator Is unable To Source " & CalibParam & CalibParamUnit
                                CanDoIt = 0
                            End If
                        Case "E"
                            If CalibParam >= -250 And CalibParam <= 1000 Then
                                CanDoIt = 1
                            Else
                                MsgBox "Calibrator Is unable To Source " & CalibParam & CalibParamUnit
                                CanDoIt = 0
                            End If
                            
                        Case "J"
                            If CalibParam >= -210 And CalibParam <= 1200 Then
                                CanDoIt = 1
                            Else
                                MsgBox "Calibrator Is unable To Source " & CalibParam & CalibParamUnit
                                CanDoIt = 0
                            End If
                            
                        Case "K"
                            If CalibParam >= -200 And CalibParam <= 1372 Then
                                CanDoIt = 1
                            Else
                                MsgBox "Calibrator Is unable To Source " & CalibParam & CalibParamUnit
                                CanDoIt = 0
                            End If
                            
                        Case "L"
                            If CalibParam >= -200 And CalibParam <= 900 Then
                                CanDoIt = 1
                            Else
                                MsgBox "Calibrator Is unable To Source " & CalibParam & CalibParamUnit
                                CanDoIt = 0
                            End If
                            
                        Case "N"
                            If CalibParam >= -200 And CalibParam <= 1300 Then
                                CanDoIt = 1
                            Else
                                MsgBox "Calibrator Is unable To Source " & CalibParam & CalibParamUnit
                                CanDoIt = 0
                            End If
                            
                        Case "R"
                            If CalibParam >= 0 And CalibParam <= 1767 Then
                                CanDoIt = 1
                            Else
                                MsgBox "Calibrator Is unable To Source " & CalibParam & CalibParamUnit
                                CanDoIt = 0
                            End If
                            
                        Case "S"
                            If CalibParam >= 0 And CalibParam <= 1767 Then
                                CanDoIt = 1
                            Else
                                MsgBox "Calibrator Is unable To Source " & CalibParam & CalibParamUnit
                                CanDoIt = 0
                            End If
                            
                        Case "T"
                            If CalibParam >= -250 And CalibParam <= 400 Then
                                CanDoIt = 1
                            Else
                                MsgBox "Calibrator Is unable To Source " & CalibParam & CalibParamUnit
                                CanDoIt = 0
                            End If
                            
                        Case "U"
                            If CalibParam >= -200 And CalibParam <= 600 Then
                                CanDoIt = 1
                            Else
                                MsgBox "Calibrator Is unable To Source " & CalibParam & CalibParamUnit
                                CanDoIt = 0
                            End If
                            
                    End Select
'------------------------End Temp Check in C Degs--------------------------------
                
                Case "Cap"
'------------------------Begin Cap Check------------------------------
                    
                    Select Case CalibParamUnit
                        Case "pF"
                            If CalibParam >= 220 And CalibParam <= 999.9999 Then
                                CanDoIt = 1
                            Else
                                
                                CalibratorSpecs "Cap", CalibParam / 1000, "nF", "", ""
                            End If
                            
                        Case "nF"
                            If CalibParam >= 1 And CalibParam <= 999.9999 Then
                                CanDoIt = 1
                            Else
                                
                                CalibratorSpecs "Cap", CalibParam / 1000, "uF", "", ""
                            End If
                        Case "uF"
                            
                            If CalibParam >= 1 And CalibParam <= 999.999 Then
                                CanDoIt = 1
                            Else
                                CalibratorSpecs "Cap", CalibParam / 1000, "mF", "", ""
                                
                            End If
                            
                        Case "mF"
                            
                            If CalibParam >= 1 And CalibParam <= 110 Then
                                CanDoIt = 1
                                
                            Else
                                MsgBox "Calibrator Is unable To Source " & CalibParam & CalibParamUnit
                                CanDoIt = 0
                            End If
                            
                    End Select
'------------------------End Cap Check--------------------------------
                    
            End Select
            
        Case "M3001"
            Select Case CalFunc
                Case "DCV"
                    '------------------------Begin DCV Check------------------------------
                    
                    If CalibParam >= 0 And CalibParam <= 100 Then
                        CanDoIt = 1
                    Else
                        MsgBox "Calibrator Is unable To Source " & CalibParam & CalibParamUnit
                        CanDoIt = 0
                    End If
                    '------------------------End DCV Check--------------------------------
                    
                Case "DCA"
                    
                    If CalibParam >= 0 And CalibParam <= 100 Then
                        CanDoIt = 1
                    Else
                        MsgBox "Calibrator Is unable To Source " & CalibParam & CalibParamUnit
                    End If
                    
            End Select
    End Select
End Sub
