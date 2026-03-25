Attribute VB_Name = "DMMSpecifications"
Sub DMMSpecs(Mode As String, CalFunc As String, CalArg1 As String, CalArg2 As String)
'Counter   "Measure",        "Freq",             "1",                ""

DMMMode = Mode
DMMCalFunc = CalFunc
DMMParam = Channel
DMMParamUnit = CalArg2


CalibratorModel = wsInfo.Range("M9").Value
CalibratorGPIB = wsInfo.Range("M11").Value
CalibratorScopeOption = wsInfo.Range("M12").Value
DMMModel = wsInfo.Range("P9").Value
DMMGPIB = wsInfo.Range("P11").Value
CounterModel = wsInfo.Range("M16").Value
CounterGPIB = wsInfo.Range("M18").Value
    
'Exit sub if no gpib address
If DMMGPIB = "" Then Exit Sub
    Set ioMgr = New VisaComLib.ResourceManager
    Set DMMDevice = New VisaComLib.FormattedIO488
    Set DMMDevice.IO = ioMgr.Open(DMMGPIB)

Select Case DMMModel

    Case "3458A"
        Select Case CalFunc
            Case "End"
'------------------------Begin DDM End Check------------------------------
                Select Case "CalArg1"
                    Case "OFF", "ON", "ALWAYS"
                        CanDoIt = 1
                    Case Else
                        MsgBox CalArg1 & " is a unknown argument for END!"
'------------------------End DMM End Check--------------------------------
                         
                End Select
                    
        End Select
End Select

 


End Sub
