Attribute VB_Name = "Control"

Sub Out()
   
    Dim HVImageLoc As String
    'instrument.WriteString "Rangelck?"
     '       CalibStatus = inst3458.ReadString()
      '      MsgBox CalibStatus
    'Wave = "square"
   'MsgBox Wave
   ' Comm True, False, False
'MsgBox OffValueV & OffValueU & OffValueHz & OffValueHzU & Wave & OffSetU & duty
    If WorkOrderSheet.Range("Calibrator") = "" Then
        
    Else
           instrument.WriteString "*cls"
           instrument.WriteString "OUT " & OffValueV & " " & OffValueU & ", " & OffValueHz & " " & OffValueHzU & "; OPER"
           If Wave <> "" Then
           instrument.WriteString "WAVE " & Wave
           End If
           If offset <> 0 Then
           instrument.WriteString "DC_OFFSET " & offset & " " & OffSetU
           End If
           If duty <> 0 Then
           instrument.WriteString "DUTY " & duty
           End If
           
           
        
    End If
    
End Sub





Sub temp(OffValueV As Double, OffValueU As String, tc_type_unit As String, comp As String)        'Use this format on sheet temp 0, "cel", "k" or temp 100, "far", "j", "Wire2" or Wire4 or None
   Dim TCValue As Integer
   Dim RTDTypeUnit As String
    
    Comm True, False, False
    
    If WorkOrderSheet.Range("Calibrator") = "" Then
      
    Else
    instrument.WriteString "*cls"
    If tc_type_unit = "ohm" Then
       ' ZCOMP Wire2 Or Wire4 Or None
        RTDTypeUnit = "PT385"
        instrument.WriteString "TSENS_TYPE" & " RTD"
        instrument.WriteString "RTD_type" & " " & RTDTypeUnit
        instrument.WriteString "ZComp" & " " & comp
    
        instrument.WriteString "OUT " & OffValueV & " " & OffValueU & "; OPER"
    Else
        instrument.WriteString "TSENS_TYPE" & " TC"
        instrument.WriteString "tc_type" & " " & tc_type_unit
        instrument.WriteString "OUT " & OffValueV & " " & OffValueU & "; OPER"
    End If
    
    End If
    
End Sub

'Sub TCMeas(temp    As Double, TUnit As String)
Sub TCMeas(OffValueU As String, AllReady As String, Source As Boolean)
    Dim inst_value  As String
    Dim inst_valuemV  As String
    Dim scientificNumber As String
    Dim decimalNumber As Double
    Dim decimalNumbermV As Double
    Dim cleanedNumber As String
    Dim cell        As Range
    Dim tc_type_unit As String
    Dim activeCol   As Long
 
    Comm True, False, False
    
    
    If WorkOrderSheet.Range("Calibrator") = "" Then
       VValmV = "53.8995E-03"
        
        'Get real number from VValmV and times by 1000 to get regular mV
        scientificNumber = VValmV
        cleanedNumber = ExtractScientificNumber(scientificNumber)
        VValmVV = CDbl(cleanedNumber) * 1000
       
        TCRefInt = "INT,2.388E+01,CEL"
        
        Else
        
        If AllReady = "Y" Then
        'TC_MEAS;*WAI;VAL?
        instrument.WriteString "tc_meas;*WAI;TC_Ref?"
        
            'Get Internal Ref Temp
           ' instrument.WriteString "TC_Ref?"
            TCRefInt = instrument.ReadString()
            
            
            'Get Internal Ref Temp as mV
            instrument.WriteString "VVal?"
            VValmV = instrument.ReadString()
            
            'Get Temp Reading
            instrument.WriteString "Val?"
            inst_value = instrument.ReadString()
            
            
            
            
        ElseIf AllReady = "n" Then
        
            'Get Internal Ref Temp
            instrument.WriteString "tc_meas;*WAI;TC_Ref?"
            TCRefInt = instrument.ReadString()
            
            
            'Get Internal Ref Temp as mV
            instrument.WriteString "VVal?"
            VValmV = instrument.ReadString()
           
            
            'Get Temp Reading
            instrument.WriteString "tc_type" & " " & OffValueU
            instrument.WriteString "Val?"
            inst_value = instrument.ReadString()
            
        End If
        
        'Get real number from inst_value
        scientificNumber = inst_value
        cleanedNumber = ExtractScientificNumber(scientificNumber)
        On Error GoTo ConversionError
        decimalNumber = CDbl(cleanedNumber)
        On Error GoTo 0
        
        
        'Get real number from VValmV and times by 1000 to get regular mV
        scientificNumber = VValmV
        cleanedNumber = ExtractScientificNumber(scientificNumber)
        On Error GoTo ConversionError
        VValmVV = CDbl(cleanedNumber) * 1000
        On Error GoTo 0
        
        If Source = True Then
        activeCol = ActiveCell.Column
        If activeCol = 6 Then
            ActiveCell = decimalNumber
        ElseIf activeCol = 7 Then
            ActiveCell = decimalNumber
        ElseIf activeCol = 8 Then
            ActiveCell = decimalNumber
        End If
        Else
        End If
        
        Exit Sub
        
ConversionError:
        MsgBox "Error: The value in cell B1 Is Not a valid scientific notation number.", vbCritical
        On Error GoTo 0
        
    End If
End Sub

Function ExtractScientificNumber(inputString As String) As String
    Dim regex       As Object
    Set regex = CreateObject("VBScript.RegExp")
    
    With regex
        .Pattern = "[-+]?[0-9]*\.?[0-9]+([eE][-+]?[0-9]+)?"
        .Global = True
    End With
    
    If regex.Test(inputString) Then
        ExtractScientificNumber = regex.Execute(inputString)(0)
    Else
        ExtractScientificNumber = "0"
    End If
End Function
