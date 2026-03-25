Attribute VB_Name = "Readings3458"


Sub Readings3458A()
    
   
    Dim inst_value  As String
    Dim err         As String
    
    Comm True, False, False
    
    'Set add3458value = Worksheets("Information").Range("DMM")
    'Set ioMgr = New VisaComLib.resourceManager
    'Set inst3458 = New VisaComLib.FormattedIO488
    'Set inst3458.IO = ioMgr.Open(add3458value.Value)
    
    If Worksheets("Information").Range("DMM") = "" Then
    
        Else
            inst3458.WriteString "func " & func
            inst3458.WriteString "END ALWAYS"
            inst3458.WriteString "NRDGS 1"
            inst3458.WriteString "TARM SGL"
            Application.Wait (Now + TimeValue("0:00:3"))
            inst_value = inst3458.ReadString()
            ' MsgBox inst_value
            inst_value2 = Left(inst_value, Len(inst_value) - 6)        ' lose last 4 digits, i.e. E-01
            inst_valueF = CDbl(inst_value2)        ' convert to float, " 5.000670612" to 5.000670612 or "-5.000670612" to -5.000670612
            inst_valueFinal = inst_valueF
             GoodRead = False
    
    'funcs   dci   dcv   diodedci

        If InStr(inst_value, "E+00") <> 0 Then    ' 1V
            inst_valueFinal = inst_valueF
           ' MsgBox "E+00"
            GoodRead = True
        End If
    
        If InStr(inst_value, "E+01") <> 0 Then    ' 10V
            inst_valueFinal = inst_valueF * 10
            'MsgBox "E+01"
            GoodRead = True
        End If
    
        If InStr(inst_value, "E+02") <> 0 Then    ' 100V
            inst_valueFinal = inst_valueF * 100
            'MsgBox "E+02"
            GoodRead = True
        End If
    
        If InStr(inst_value, "E+03") <> 0 Then    ' 1000V
            inst_valueFinal = inst_valueF * 1000
            'MsgBox "E+03"
            GoodRead = True
        End If
    
        If InStr(inst_value, "E+04") <> 0 Then    ' 10000V
            inst_valueFinal = inst_valueF * 10000
           ' MsgBox inst_valueFinal
            inst_valueFinal = inst_valueF
            'MsgBox inst_valueFinal
            GoodRead = True
        End If
        
        If InStr(inst_value, "E-10") <> 0 Then    ' 0.0000001mV
            inst_valueFinal = inst_valueF / 10000000000#
            'MsgBox "E-10"
            GoodRead = True
        End If
    
        If InStr(inst_value, "E-09") <> 0 Then    ' 0.000001mV
            inst_valueFinal = inst_valueF / 1000000000
            'MsgBox "E-09"
            GoodRead = True
        End If
    
        If InStr(inst_value, "E-08") <> 0 Then    ' 0.00001mV
            inst_valueFinal = inst_valueF / 100000000
            'MsgBox "E-08"
            GoodRead = True
        End If
    
        If InStr(inst_value, "E-07") <> 0 Then    ' 0.0001mV
            inst_valueFinal = inst_valueF / 10000000
            'MsgBox "E-07"
            GoodRead = True
        End If
    
        If InStr(inst_value, "E-06") <> 0 Then    ' 0.001mV
            inst_valueFinal = inst_valueF / 1000000
            'MsgBox "E-06"
            GoodRead = True
        End If
    
        If InStr(inst_value, "E-05") <> 0 Then    ' 0.01mV
            inst_valueFinal = inst_valueF / 100000
            'MsgBox "E-05"
            GoodRead = True
        End If
    
        If InStr(inst_value, "E-04") <> 0 Then    ' 0.1mV uA
            inst_valueFinal = inst_valueF / 10000
            'MsgBox "E-04"
            GoodRead = True
        End If
        
        If InStr(inst_value, "E-03") <> 0 Then '1 mA
           ' MsgBox inst_valueF
            inst_valueFinal = inst_valueF / 1
           ' MsgBox "e-03"
            GoodRead = True
        End If
        
        If InStr(inst_value, "E-02") <> 0 Then    ' 100 mA
            inst_valueFinal = inst_valueF * 10
           ' MsgBox "E-02"
            GoodRead = True
        End If
        
        If InStr(inst_value, "E-01") <> 0 Then    ' 100mV
            inst_valueFinal = inst_valueF / 10
            'MsgBox "E-01"
            GoodRead = True
        End If


    
    'End If
    
    ' Check if successful reading received
    'If GoodRead = True Then
       
     '   activeCell = inst_valueFinal  'update spreadsheet
     
   ' End If
    
    'If GoodRead = False Then
    '    Count = Count - 1           ' bad read so re-read
    'End If

  

    ' Set the value of the active cell
    'ActiveCell.value = inst_valueF

   ' Application.Wait (Now + TimeValue("0:00:3"))
  ' Call controlrencmd
    'instrument.IO.Close

        End If
    

    
    
    Exit Sub ' Exit gracefully if no error occurs
End Sub


