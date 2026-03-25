Attribute VB_Name = "DatasheetCode"

Public Sub HandleSelectionChange(ByVal Target As Excel.Range)
    'Load PanelForm
    AutoSelect = True
    DiffTitle = ""
    ShowPanel
    On Error Resume Next
    AppActivate Application.Caption
    On Error GoTo 0
    On Error GoTo ErrorHandler
    Dim i           As Long
    Dim startRow    As Long
    Dim endRow      As Long
    Dim currentRow  As Long
    Dim StandOper   As Boolean
    Dim volt        As Double
    
    CalibratorModel = wsInfo.Range("M9").Value
    DMMModel = wsInfo.Range("P9").Value
    CounterModel = wsInfo.Range("M16").Value
    SetupWS
    '-------------------Begin Reassigning Standard GPIB addresses-------------------
    CalibratorGPIB = wsInfo.Range("M11").Value
    DMMGPIB = wsInfo.Range("P11").Value
    CounterGPIB = wsInfo.Range("M18").Value
    '-------------------End Reassigning Standard GPIB addresses---------------------
    
    '-------------------Begin Standard Reset Check----------------------------------
    ' === Calibrator Clear using Reset===
    HVImageShow 0, ""
    If CalibratorGPIB = "" Then
        
        If CalibratorReset = 1 Then
            
        Else
            
            CalibClearStatus "Clear"
            CalibClearStatus "Standby"
            CalibratorReset = 1
            
        End If
    End If
    
    ' === DMM Clear using Reset===
    'If DMMGPIB = "" Then
    
    '   If DMMReset = 1 Then
    
    '  Else
    '    MsgBox "Reset?"
    '     DMMClearStatus "Reset"
    '   DMMReset = 1
    
    'End If
    'End If
    
    ' === Counter Clear using Reset===
    'If CounterGPIB = "" Then
        
     '   If CounterReset = 1 Then
            
      '  Else
            
       '     CounterClearStatus "Clear"
        '    CounterReset = 1
            
        'End If
    'End If
    '-------------------End Standard Reset Check----------------------------------
    
'---------------------------------------------------------------------------------
'---------------------------------Edit Below--------------------------------------
'---------------------------------------------------------------------------------

If ToggleStates("CodeButton") = "Off" Then
   
Else
'This is the command line for source with calibrator. just copy and paste and make changes to what you need.

'             Mode,  CalFunc, param, paramUnit, Hertz, HertzUnit, Wave, OffSet, Duty, ZComp
'Calibrator "Source", "Ohm",   350,    "ohm",     0,      "",      "",     0,     0,  "Wire4"

'If you are doing DC Voltage at 15 V -------- make sure anything with letters gets quotes around it ""
'Calibrator "Source", "DCV", 15, "V", 0, "Hz", "", 0, 0, ""
'Doing 4 wire resistance 350 ohms  or Wire2 or none
'Calibrator "Source", "Ohm", 350, "ohm", 0, "", "", 0, 0, "Wire4"
        
        '-------------------Begin All Test Points on Datasheet---------------------
        If Target Is Nothing Then Exit Sub
        Select Case Target.Address
            
            Case "$F$13", "$G$13"
'AC Voltage Tests @ 60 Hz
                
                TestSections 1
                '-------------------Begin Shows the MainHookup Userform for hookups and pictures---------------------
                If PrevTestSect <> TestSect Then UForms "MainForm"
                If TerminateClicked Then TerminateClicked = False: Exit Sub
                '-------------------End Shows the MainHookup Userform for hookups and pictures-----------------------
                
                Calibrator "Source", "ACV", 100, "mV", 60, "Hz", "", 0, 0, ""
                PrevTestSect = TestSect
                
            Case "$F$14", "$G$14"
                
                TestSections 1
                '-------------------Begin Shows the MainHookup Userform for hookups and pictures---------------------
                If PrevTestSect <> TestSect Then UForms "MainForm"
                If TerminateClicked Then TerminateClicked = False: Exit Sub
                '-------------------End Shows the MainHookup Userform for hookups and pictures-----------------------
                
                Calibrator "Source", "ACV", 300, "mV", 60, "Hz", "", 0, 0, ""
                PrevTestSect = TestSect
                
            Case "$F$15", "$G$15"
                Selection.OffSet(1, 0).Select
                
            Case "$F$16", "$G$16"
            
                TestSections 1
                '-------------------Begin Shows the MainHookup Userform for hookups and pictures---------------------
                If PrevTestSect <> TestSect Then UForms "MainForm"
                If TerminateClicked Then TerminateClicked = False: Exit Sub
                '-------------------End Shows the MainHookup Userform for hookups and pictures-----------------------
                
                Calibrator "Source", "ACV", 1, "V", 60, "Hz", "", 0, 0, ""
                PrevTestSect = TestSect
                
            Case "$F$17", "$G$17"
            
                TestSections 1
                '-------------------Begin Shows the MainHookup Userform for hookups and pictures---------------------
                If PrevTestSect <> TestSect Then UForms "MainForm"
                If TerminateClicked Then TerminateClicked = False: Exit Sub
                '-------------------End Shows the MainHookup Userform for hookups and pictures-----------------------
                
                Calibrator "Source", "ACV", 2, "V", 60, "Hz", "", 0, 0, ""
                PrevTestSect = TestSect
                
            Case "$F$18", "$G$18"
            
                TestSections 1
                '-------------------Begin Shows the MainHookup Userform for hookups and pictures---------------------
                If PrevTestSect <> TestSect Then UForms "MainForm"
                If TerminateClicked Then TerminateClicked = False: Exit Sub
                '-------------------End Shows the MainHookup Userform for hookups and pictures-----------------------
                
                Calibrator "Source", "ACV", 3, "V", 60, "Hz", "", 0, 0, ""
                PrevTestSect = TestSect
                
            Case "$F$19", "$G$19"
            
                TestSections 1
                '-------------------Begin Shows the MainHookup Userform for hookups and pictures---------------------
                If PrevTestSect <> TestSect Then UForms "MainForm"
                If TerminateClicked Then TerminateClicked = False: Exit Sub
                '-------------------End Shows the MainHookup Userform for hookups and pictures-----------------------
                
                Calibrator "Source", "ACV", 10, "V", 60, "Hz", "", 0, 0, ""
                PrevTestSect = TestSect
                
            Case "$F$20", "$G$20"
            
                TestSections 1
                '-------------------Begin Shows the MainHookup Userform for hookups and pictures---------------------
                If PrevTestSect <> TestSect Then UForms "MainForm"
                If TerminateClicked Then TerminateClicked = False: Exit Sub
                '-------------------End Shows the MainHookup Userform for hookups and pictures-----------------------
                
                Calibrator "Source", "ACV", 30, "V", 60, "Hz", "", 0, 0, ""

                PrevTestSect = TestSect
                
            Case "$F$21", "$G$21"
                
                TestSections 1
                '-------------------Begin Shows the MainHookup Userform for hookups and pictures---------------------
                If PrevTestSect <> TestSect Then UForms "MainForm"
                If TerminateClicked Then TerminateClicked = False: Exit Sub
                '-------------------End Shows the MainHookup Userform for hookups and pictures-----------------------
                
                Calibrator "Source", "ACV", 100, "V", 60, "Hz", "", 0, 0, ""

                PrevTestSect = TestSect
                
            Case "$F$22", "$G$22"
                
                TestSections 1
                '-------------------Begin Shows the MainHookup Userform for hookups and pictures---------------------
                If PrevTestSect <> TestSect Then UForms "MainForm"
                If TerminateClicked Then TerminateClicked = False: Exit Sub
                '-------------------End Shows the MainHookup Userform for hookups and pictures-----------------------
                
                Calibrator "Source", "ACV", 300, "V", 60, "Hz", "", 0, 0, ""

                PrevTestSect = TestSect
                
            Case "$F$23", "$G$23"
                
                TestSections 2
                '-------------------Begin Shows the MainHookup Userform for hookups and pictures---------------------
                If PrevTestSect <> TestSect Then UForms "MainForm"
                If TerminateClicked Then TerminateClicked = False: Exit Sub
                '-------------------End Shows the MainHookup Userform for hookups and pictures-----------------------
                
                Calibrator "Source", "ACV", 100, "V", 60, "Hz", "", 0, 0, ""

                PrevTestSect = TestSect
                
            Case "$F$24", "$G$24"
                
                TestSections 2
                '-------------------Begin Shows the MainHookup Userform for hookups and pictures---------------------
                If PrevTestSect <> TestSect Then UForms "MainForm"
                If TerminateClicked Then TerminateClicked = False: Exit Sub
                '-------------------End Shows the MainHookup Userform for hookups and pictures-----------------------
                
                Calibrator "Source", "ACV", 800, "V", 60, "Hz", "", 0, 0, ""

                PrevTestSect = TestSect
                
            Case "$F$25", "$G$25"
                CalibClearStatus "Reset"
                Selection.OffSet(1, 0).Select
                
            Case "$F$26", "$G$26"
                Selection.OffSet(1, 0).Select
                

'Frequency Tests @ 5 Vrms
            Case "$F$27", "$G$27"
                
                TestSections 3
                '-------------------Begin Shows the MainHookup Userform for hookups and pictures---------------------
                If PrevTestSect <> TestSect Then UForms "MainForm"
                If TerminateClicked Then TerminateClicked = False: Exit Sub
                '-------------------End Shows the MainHookup Userform for hookups and pictures-----------------------
                
                Calibrator "Source", "ACV", 5, "V", 100, "Hz", "", 0, 0, ""

                PrevTestSect = TestSect
                
            Case "$F$28", "$G$28"
                
                TestSections 3
                '-------------------Begin Shows the MainHookup Userform for hookups and pictures---------------------
                If PrevTestSect <> TestSect Then UForms "MainForm"
                If TerminateClicked Then TerminateClicked = False: Exit Sub
                '-------------------End Shows the MainHookup Userform for hookups and pictures-----------------------
                
                Calibrator "Source", "ACV", 5, "V", 1, "kHz", "", 0, 0, ""

                PrevTestSect = TestSect
                
            Case "$F$29", "$G$29"
                
                TestSections 3
                '-------------------Begin Shows the MainHookup Userform for hookups and pictures---------------------
                If PrevTestSect <> TestSect Then UForms "MainForm"
                If TerminateClicked Then TerminateClicked = False: Exit Sub
                '-------------------End Shows the MainHookup Userform for hookups and pictures-----------------------
                
                Calibrator "Source", "ACV", 5, "V", 10, "kHz", "", 0, 0, ""

                PrevTestSect = TestSect
                
            Case "$F$30", "$G$30"
                CalibClearStatus "Reset"
                Selection.OffSet(1, 0).Select
                
            Case "$F$31", "$G$31"
                Selection.OffSet(1, 0).Select


'DC Voltage Tests
            Case "$F$32", "$G$32"
                
                TestSections 4
                '-------------------Begin Shows the MainHookup Userform for hookups and pictures---------------------
                If PrevTestSect <> TestSect Then UForms "MainForm"
                If TerminateClicked Then TerminateClicked = False: Exit Sub
                '-------------------End Shows the MainHookup Userform for hookups and pictures-----------------------
                
                Calibrator "Source", "DCV", 1, "V", 0, "Hz", "", 0, 0, ""
                PrevTestSect = TestSect
                
            Case "$F$33", "$G$33"
                
                TestSections 4
                '-------------------Begin Shows the MainHookup Userform for hookups and pictures---------------------
                If PrevTestSect <> TestSect Then UForms "MainForm"
                If TerminateClicked Then TerminateClicked = False: Exit Sub
                '-------------------End Shows the MainHookup Userform for hookups and pictures-----------------------
                
                Calibrator "Source", "DCV", 3, "V", 0, "Hz", "", 0, 0, ""
                PrevTestSect = TestSect
                
            Case "$F$34", "$G$34"
                
                TestSections 4
                '-------------------Begin Shows the MainHookup Userform for hookups and pictures---------------------
                If PrevTestSect <> TestSect Then UForms "MainForm"
                If TerminateClicked Then TerminateClicked = False: Exit Sub
                '-------------------End Shows the MainHookup Userform for hookups and pictures-----------------------
                
                Calibrator "Source", "DCV", 10, "V", 0, "Hz", "", 0, 0, ""
                PrevTestSect = TestSect
                
            Case "$F$35", "$G$35"
                
                TestSections 4
                '-------------------Begin Shows the MainHookup Userform for hookups and pictures---------------------
                If PrevTestSect <> TestSect Then UForms "MainForm"
                If TerminateClicked Then TerminateClicked = False: Exit Sub
                '-------------------End Shows the MainHookup Userform for hookups and pictures-----------------------
                
                Calibrator "Source", "DCV", 30, "V", 0, "Hz", "", 0, 0, ""
                PrevTestSect = TestSect
                
            Case "$F$36", "$G$36"
                
                TestSections 4
                '-------------------Begin Shows the MainHookup Userform for hookups and pictures---------------------
                If PrevTestSect <> TestSect Then UForms "MainForm"
                If TerminateClicked Then TerminateClicked = False: Exit Sub
                '-------------------End Shows the MainHookup Userform for hookups and pictures-----------------------
                
                Calibrator "Source", "DCV", 100, "V", 0, "Hz", "", 0, 0, ""
                PrevTestSect = TestSect
                
            Case "$F$37", "$G$37"
                
                TestSections 4
                '-------------------Begin Shows the MainHookup Userform for hookups and pictures---------------------
                If PrevTestSect <> TestSect Then UForms "MainForm"
                If TerminateClicked Then TerminateClicked = False: Exit Sub
                '-------------------End Shows the MainHookup Userform for hookups and pictures-----------------------
                
                Calibrator "Source", "DCV", 300, "V", 0, "Hz", "", 0, 0, ""
                PrevTestSect = TestSect
                
            Case "$F$38", "$G$38"
                
                TestSections 5
                '-------------------Begin Shows the MainHookup Userform for hookups and pictures---------------------
                If PrevTestSect <> TestSect Then UForms "MainForm"
                If TerminateClicked Then TerminateClicked = False: Exit Sub
                '-------------------End Shows the MainHookup Userform for hookups and pictures-----------------------
                
                Calibrator "Source", "DCV", 100, "V", 0, "Hz", "", 0, 0, ""
                PrevTestSect = TestSect
                
            Case "$F$39", "$G$39"
                
                TestSections 5
                '-------------------Begin Shows the MainHookup Userform for hookups and pictures---------------------
                If PrevTestSect <> TestSect Then UForms "MainForm"
                If TerminateClicked Then TerminateClicked = False: Exit Sub
                '-------------------End Shows the MainHookup Userform for hookups and pictures-----------------------
                
                Calibrator "Source", "DCV", 800, "V", 0, "Hz", "", 0, 0, ""
                PrevTestSect = TestSect

                
            Case "$F$40", "$G$40"
            CalibClearStatus "Reset"
                Selection.OffSet(1, 0).Select
                
'DC Millivolt Tests
                
            Case "$F$41", "$G$41"
                
                TestSections 6
                '-------------------Begin Shows the MainHookup Userform for hookups and pictures---------------------
                If PrevTestSect <> TestSect Then UForms "MainForm"
                If TerminateClicked Then TerminateClicked = False: Exit Sub
                '-------------------End Shows the MainHookup Userform for hookups and pictures-----------------------
                
                Calibrator "Source", "DCV", 100, "mV", 0, "Hz", "", 0, 0, ""
                PrevTestSect = TestSect
                
            Case "$F$42", "$G$42"
                
                TestSections 6
                '-------------------Begin Shows the MainHookup Userform for hookups and pictures---------------------
                If PrevTestSect <> TestSect Then UForms "MainForm"
                If TerminateClicked Then TerminateClicked = False: Exit Sub
                '-------------------End Shows the MainHookup Userform for hookups and pictures-----------------------
                
                Calibrator "Source", "DCV", 300, "mV", 0, "Hz", "", 0, 0, ""
                PrevTestSect = TestSect
                
            Case "$F$43", "$G$43"
                CalibClearStatus "Reset"
                Selection.OffSet(1, 0).Select
                
            Case "$F$44", "$G$44"
                Selection.OffSet(1, 0).Select
                
 'Resistance Tests
            Case "$F$45", "$G$45"
            
                TestSections 7
                '-------------------Begin Shows the MainHookup Userform for hookups and pictures---------------------
                If PrevTestSect <> TestSect Then UForms "MainForm"
                If TerminateClicked Then TerminateClicked = False: Exit Sub
                '-------------------End Shows the MainHookup Userform for hookups and pictures-----------------------
                
                Calibrator "Source", "Ohm", 120, "ohm", 0, "", "", 0, 0, "Wire2"
                PrevTestSect = TestSect

            Case "$F$46", "$G$46"
            
                TestSections 7
                '-------------------Begin Shows the MainHookup Userform for hookups and pictures---------------------
                If PrevTestSect <> TestSect Then UForms "MainForm"
                If TerminateClicked Then TerminateClicked = False: Exit Sub
                '-------------------End Shows the MainHookup Userform for hookups and pictures-----------------------
                
                Calibrator "Source", "Ohm", 300, "ohm", 0, "", "", 0, 0, "Wire2"
                PrevTestSect = TestSect
                
            Case "$F$47", "$G$47"
                Selection.OffSet(1, 0).Select

            Case "$F$48", "$G$48"
            
                TestSections 7
                '-------------------Begin Shows the MainHookup Userform for hookups and pictures---------------------
                If PrevTestSect <> TestSect Then UForms "MainForm"
                If TerminateClicked Then TerminateClicked = False: Exit Sub
                '-------------------End Shows the MainHookup Userform for hookups and pictures-----------------------
                
                Calibrator "Source", "Ohm", 1.2, "kohm", 0, "", "", 0, 0, "Wire2"
                PrevTestSect = TestSect

            Case "$F$49", "$G$49"
            
                TestSections 7
                '-------------------Begin Shows the MainHookup Userform for hookups and pictures---------------------
                If PrevTestSect <> TestSect Then UForms "MainForm"
                If TerminateClicked Then TerminateClicked = False: Exit Sub
                '-------------------End Shows the MainHookup Userform for hookups and pictures-----------------------
                
                Calibrator "Source", "Ohm", 3, "kohm", 0, "", "", 0, 0, "Wire2"
                PrevTestSect = TestSect

            Case "$F$50", "$G$50"
            
                TestSections 7
                '-------------------Begin Shows the MainHookup Userform for hookups and pictures---------------------
                If PrevTestSect <> TestSect Then UForms "MainForm"
                If TerminateClicked Then TerminateClicked = False: Exit Sub
                '-------------------End Shows the MainHookup Userform for hookups and pictures-----------------------
                
                Calibrator "Source", "Ohm", 12, "kohm", 0, "", "", 0, 0, "Wire2"
                PrevTestSect = TestSect

            Case "$F$51", "$G$51"
            
                TestSections 7
                '-------------------Begin Shows the MainHookup Userform for hookups and pictures---------------------
                If PrevTestSect <> TestSect Then UForms "MainForm"
                If TerminateClicked Then TerminateClicked = False: Exit Sub
                '-------------------End Shows the MainHookup Userform for hookups and pictures-----------------------
                
                Calibrator "Source", "Ohm", 30, "kohm", 0, "", "", 0, 0, "Wire2"
                PrevTestSect = TestSect

            Case "$F$52", "$G$52"
            
                TestSections 7
                '-------------------Begin Shows the MainHookup Userform for hookups and pictures---------------------
                If PrevTestSect <> TestSect Then UForms "MainForm"
                If TerminateClicked Then TerminateClicked = False: Exit Sub
                '-------------------End Shows the MainHookup Userform for hookups and pictures-----------------------
                
                Calibrator "Source", "Ohm", 120, "kohm", 0, "", "", 0, 0, "None"
                PrevTestSect = TestSect

            Case "$F$53", "$G$53"
            
                TestSections 7
                '-------------------Begin Shows the MainHookup Userform for hookups and pictures---------------------
                If PrevTestSect <> TestSect Then UForms "MainForm"
                If TerminateClicked Then TerminateClicked = False: Exit Sub
                '-------------------End Shows the MainHookup Userform for hookups and pictures-----------------------
                
                Calibrator "Source", "Ohm", 200, "kohm", 0, "", "", 0, 0, "None"
                PrevTestSect = TestSect

            Case "$F$54", "$G$54"
            
                TestSections 7
                '-------------------Begin Shows the MainHookup Userform for hookups and pictures---------------------
                If PrevTestSect <> TestSect Then UForms "MainForm"
                If TerminateClicked Then TerminateClicked = False: Exit Sub
                '-------------------End Shows the MainHookup Userform for hookups and pictures-----------------------
                
                Calibrator "Source", "Ohm", 300, "kohm", 0, "", "", 0, 0, "None"
                PrevTestSect = TestSect
                
            Case "$F$55", "$G$55"
                Selection.OffSet(1, 0).Select
                

            Case "$F$56", "$G$56"
            
                TestSections 7
                '-------------------Begin Shows the MainHookup Userform for hookups and pictures---------------------
                If PrevTestSect <> TestSect Then UForms "MainForm"
                If TerminateClicked Then TerminateClicked = False: Exit Sub
                '-------------------End Shows the MainHookup Userform for hookups and pictures-----------------------
                
                Calibrator "Source", "Ohm", 1.2, "Mohm", 0, "", "", 0, 0, "None"
                PrevTestSect = TestSect
                

            Case "$F$57", "$G$57"
            
                TestSections 7
                '-------------------Begin Shows the MainHookup Userform for hookups and pictures---------------------
                If PrevTestSect <> TestSect Then UForms "MainForm"
                If TerminateClicked Then TerminateClicked = False: Exit Sub
                '-------------------End Shows the MainHookup Userform for hookups and pictures-----------------------
                
                Calibrator "Source", "Ohm", 3, "Mohm", 0, "", "", 0, 0, "None"
                PrevTestSect = TestSect
                
            Case "$F$58", "$G$58"
                Selection.OffSet(1, 0).Select
                

            Case "$F$59", "$G$59"
            
                TestSections 7
                '-------------------Begin Shows the MainHookup Userform for hookups and pictures---------------------
                If PrevTestSect <> TestSect Then UForms "MainForm"
                If TerminateClicked Then TerminateClicked = False: Exit Sub
                '-------------------End Shows the MainHookup Userform for hookups and pictures-----------------------
                
                Calibrator "Source", "Ohm", 12, "Mohm", 0, "", "", 0, 0, "None"
                PrevTestSect = TestSect
                

            Case "$F$60", "$G$60"
            
                TestSections 7
                '-------------------Begin Shows the MainHookup Userform for hookups and pictures---------------------
                If PrevTestSect <> TestSect Then UForms "MainForm"
                If TerminateClicked Then TerminateClicked = False: Exit Sub
                '-------------------End Shows the MainHookup Userform for hookups and pictures-----------------------
                
                Calibrator "Source", "Ohm", 30, "Mohm", 0, "", "", 0, 0, "None"
                PrevTestSect = TestSect
                
            Case "$F$61", "$G$61"
                CalibClearStatus "Reset"
                Selection.OffSet(1, 0).Select

'Continuity Beeper of @ 250
                

            Case "$F$62", "$G$62"
            
                TestSections 8
                DiffTitle = "B62"
                '-------------------Begin Shows the MainHookup Userform for hookups and pictures---------------------
                If PrevTestSect <> TestSect Then UForms "MainForm"
                If TerminateClicked Then TerminateClicked = False: Exit Sub
                '-------------------End Shows the MainHookup Userform for hookups and pictures-----------------------
                TestSections 10000
                VariableString = "Off at 250 ohms"
                Calibrator "Source", "Ohm", 310, "ohm", 0, "", "", 0, 0, "None"
                Application.Wait (Now + TimeValue("0:00:2"))
                Calibrator "Source", "Ohm", 250, "ohm", 0, "", "", 0, 0, "None"
                If PrevTestSect <> TestSect Then UForms "MainForm_Basic_Comment"
                If TerminateClicked Then TerminateClicked = False: Exit Sub
                TestSect = TestSectBak
                PrevTestSect = TestSect
                
'Continuity Beeper on @ 100
            Case "$F$63", "$G$63"
            
                TestSections 8
                DiffTitle = "B63"
                '-------------------Begin Shows the MainHookup Userform for hookups and pictures---------------------
                If PrevTestSect <> TestSect Then UForms "MainForm"
                If TerminateClicked Then TerminateClicked = False: Exit Sub
                '-------------------End Shows the MainHookup Userform for hookups and pictures-----------------------
                TestSections 10000
                VariableString = "On at 100 ohms"
                Calibrator "Source", "Ohm", 100, "ohm", 0, "", "", 0, 0, "None"
                If PrevTestSect <> TestSect Then UForms "MainForm_Basic_Comment"
                If TerminateClicked Then TerminateClicked = False: Exit Sub
                TestSect = TestSectBak
                PrevTestSect = TestSect
                
                

                
            Case "$F$64", "$G$64"
                CalibClearStatus "Reset"
                Selection.OffSet(1, 0).Select
                
            Case "$F$65", "$G$65"
                Selection.OffSet(1, 0).Select
                
'Diode Test
                
            Case "$F$66", "$G$66"
            
                TestSections 9
                '-------------------Begin Shows the MainHookup Userform for hookups and pictures---------------------
                If PrevTestSect <> TestSect Then UForms "MainForm"
                If TerminateClicked Then TerminateClicked = False: Exit Sub
                '-------------------End Shows the MainHookup Userform for hookups and pictures-----------------------
                
                Calibrator "Source", "DCV", 2, "V", 0, "Hz", "", 0, 0, ""
                PrevTestSect = TestSect
           
'Unit outputs 0.2-0.33
            Case "$F$67", "$G$67"
            
                TestSections 10
                '-------------------Begin Shows the MainHookup Userform for hookups and pictures---------------------
                If PrevTestSect <> TestSect Then UForms "MainForm"
                If TerminateClicked Then TerminateClicked = False: Exit Sub
                '-------------------End Shows the MainHookup Userform for hookups and pictures-----------------------
                
                    Multiplier = 1000
                    DMMClearStatus "Reset"
                    DMM "END", "ALWAYS"
                    DMM "Func", "DCI"
                    DMMReset = 1
                    DMM "Range", "1.0E-3"
                    DMM "TRIG", "SGL"
                    MsgBox FixedRdg
                    If FixedRdg >= 0.2 And FixedRdg <= 0.33 Then
                    ActiveCell.Value = "Pass"
                    Else
                    ActiveCell.Value = "Fail"
                    End If
                    Selection.OffSet(1, 0).Select
                    PrevTestSect = TestSect

                
            Case "$F$68", "$G$68"
                CalibClearStatus "Reset"
                Selection.OffSet(1, 0).Select
                
            Case "$F$69", "$G$69"
                Selection.OffSet(1, 0).Select
                
'DC Milliamps Tests
                
            Case "$F$70", "$G$70"
            
                TestSections 11
                '-------------------Begin Shows the MainHookup Userform for hookups and pictures---------------------
                If PrevTestSect <> TestSect Then UForms "MainForm"
                If TerminateClicked Then TerminateClicked = False: Exit Sub
                '-------------------End Shows the MainHookup Userform for hookups and pictures-----------------------
                
                Calibrator "Source", "DCI", 4, "mA", 0, "Hz", "", 0, 0, "None"
                PrevTestSect = TestSect
                
            Case "$F$71", "$G$71"
            
                TestSections 11
                '-------------------Begin Shows the MainHookup Userform for hookups and pictures---------------------
                If PrevTestSect <> TestSect Then UForms "MainForm"
                If TerminateClicked Then TerminateClicked = False: Exit Sub
                '-------------------End Shows the MainHookup Userform for hookups and pictures-----------------------
                
                Calibrator "Source", "DCI", 12, "mA", 0, "Hz", "", 0, 0, "None"
                PrevTestSect = TestSect
                
            Case "$F$72", "$G$72"
            
                TestSections 11
                '-------------------Begin Shows the MainHookup Userform for hookups and pictures---------------------
                If PrevTestSect <> TestSect Then UForms "MainForm"
                If TerminateClicked Then TerminateClicked = False: Exit Sub
                '-------------------End Shows the MainHookup Userform for hookups and pictures-----------------------
                
                Calibrator "Source", "DCV", 20, "mA", 0, "Hz", "", 0, 0, "None"
                PrevTestSect = TestSect

                
            Case "$F$73", "$G$73"
                CalibClearStatus "Reset"
                Selection.OffSet(1, 0).Select
                
            Case "$F$74", "$G$74"
                Selection.OffSet(1, 0).Select
                
'DC Amps Tests
                
            Case "$F$75", "$G$75"
            
                TestSections 12
                '-------------------Begin Shows the MainHookup Userform for hookups and pictures---------------------
                If PrevTestSect <> TestSect Then UForms "MainForm"
                If TerminateClicked Then TerminateClicked = False: Exit Sub
                '-------------------End Shows the MainHookup Userform for hookups and pictures-----------------------
                
                Calibrator "Source", "DCI", 0.1, "A", 0, "Hz", "", 0, 0, "None"
                PrevTestSect = TestSect
                
            Case "$F$76", "$G$76"
            
                TestSections 12
                '-------------------Begin Shows the MainHookup Userform for hookups and pictures---------------------
                If PrevTestSect <> TestSect Then UForms "MainForm"
                If TerminateClicked Then TerminateClicked = False: Exit Sub
                '-------------------End Shows the MainHookup Userform for hookups and pictures-----------------------
                
                Calibrator "Source", "DCI", 0.4, "A", 0, "Hz", "", 0, 0, "None"
                PrevTestSect = TestSect

                
            Case "$F$77", "$G$77"
                CalibClearStatus "Reset"
                Selection.OffSet(1, 0).Select
                
            Case "$F$78", "$G$78"
                Selection.OffSet(1, 0).Select
                
'AC Current Tests @ 60Hz
                
            Case "$F$79", "$G$79"
            
                TestSections 13
                '-------------------Begin Shows the MainHookup Userform for hookups and pictures---------------------
                If PrevTestSect <> TestSect Then UForms "MainForm"
                If TerminateClicked Then TerminateClicked = False: Exit Sub
                '-------------------End Shows the MainHookup Userform for hookups and pictures-----------------------
                
                Calibrator "Source", "DCI", 0.1, "A", 60, "Hz", "", 0, 0, "None"
                PrevTestSect = TestSect
                
            Case "$F$80", "$G$80"
            
                TestSections 13
                '-------------------Begin Shows the MainHookup Userform for hookups and pictures---------------------
                If PrevTestSect <> TestSect Then UForms "MainForm"
                If TerminateClicked Then TerminateClicked = False: Exit Sub
                '-------------------End Shows the MainHookup Userform for hookups and pictures-----------------------
                
                Calibrator "Source", "DCI", 0.4, "A", 60, "Hz", "", 0, 0, "None"
                PrevTestSect = TestSect

                
            Case "$F$81", "$G$81"
                CalibClearStatus "Reset"
                Selection.OffSet(1, 0).Select
                
            Case "$F$82", "$G$82"
                Selection.OffSet(1, 0).Select
           
'Current Source Tests

            Case "$F$83", "$G$83"
            
                TestSections 14
                '-------------------Begin Shows the MainHookup Userform for hookups and pictures---------------------
                If PrevTestSect <> TestSect Then UForms "MainForm"
                If TerminateClicked Then TerminateClicked = False: Exit Sub
                '-------------------End Shows the MainHookup Userform for hookups and pictures-----------------------
                TestSections 6000
                VariableString = "4 mA"
                If PrevTestSect <> TestSect Then UForms "MainForm_Basic_Comment"
                If TerminateClicked Then TerminateClicked = False: Exit Sub
                

                    '===========================
                    Multiplier = 1000
                    DMMClearStatus "Reset"
                    DMM "END", "ALWAYS"
                    DMM "Func", "DCI"
                    DMMReset = 1
                    DMM "Range", "10E-3"
                    DMM "TRIG", "SGL"
                    ActiveCell = FixedRdg
                    TestSect = TestSectBak
                    PrevTestSect = TestSect

            Case "$F$84", "$G$84"
            
                TestSections 14
                '-------------------Begin Shows the MainHookup Userform for hookups and pictures---------------------
                If PrevTestSect <> TestSect Then UForms "MainForm"
                If TerminateClicked Then TerminateClicked = False: Exit Sub
                '-------------------End Shows the MainHookup Userform for hookups and pictures-----------------------
                    TestSections 6000
                    VariableString = "12 mA"
                    If PrevTestSect <> TestSect Then UForms "MainForm_Basic_Comment"
                    If TerminateClicked Then TerminateClicked = False: Exit Sub
                    '===========================
                    Multiplier = 1000
                    DMMClearStatus "Reset"
                    DMM "END", "ALWAYS"
                    DMM "Func", "DCI"
                    DMMReset = 1
                    DMM "Range", "100E-3"
                    DMM "TRIG", "SGL"
                    ActiveCell = FixedRdg
                    TestSect = TestSectBak
                    PrevTestSect = TestSect

            Case "$F$85", "$G$85"
            
                TestSections 14
                '-------------------Begin Shows the MainHookup Userform for hookups and pictures---------------------
                If PrevTestSect <> TestSect Then UForms "MainForm"
                If TerminateClicked Then TerminateClicked = False: Exit Sub
                '-------------------End Shows the MainHookup Userform for hookups and pictures-----------------------
                    TestSections 6000
                    VariableString = "20 mA"
                    If PrevTestSect <> TestSect Then UForms "MainForm_Basic_Comment"
                    If TerminateClicked Then TerminateClicked = False: Exit Sub
                    '===========================
                    Multiplier = 1000
                    DMMClearStatus "Reset"
                    DMM "END", "ALWAYS"
                    DMM "Func", "DCI"
                    DMMReset = 1
                    DMM "Range", "100E-3"
                    DMM "TRIG", "SGL"
                    ActiveCell = FixedRdg
                    TestSect = TestSectBak
                    PrevTestSect = TestSect

                
            Case "$F$86", "$G$86"
                CalibClearStatus "Reset"
                Selection.OffSet(1, 0).Select
                
            Case "$F$87", "$G$87"
                Selection.OffSet(1, 0).Select
                
'Open Circuit Voltage

            Case "$F$88", "$G$88"
            
                TestSections 15
                '-------------------Begin Shows the MainHookup Userform for hookups and pictures---------------------
                If PrevTestSect <> TestSect Then UForms "MainForm"
                If TerminateClicked Then TerminateClicked = False: Exit Sub
                '-------------------End Shows the MainHookup Userform for hookups and pictures-----------------------
                    If DMMGPIB = "" Then
                    TestSections 9000
                    SectTitleString = "B88"
                    CommentsString = "Is the Output voltage from source jacks between 29.8 to 32V?"
                    If PrevTestSect <> TestSect Then UForms "MainForm_Basic_Comment"
                    If TerminateClicked Then TerminateClicked = False: Exit Sub
                    Else
                    AutoSelect = False
                    Multiplier = 1
                    DMMClearStatus "Reset"
                    DMM "END", "ALWAYS"
                    DMM "Func", "DCV"
                    DMMReset = 1
                    DMM "Range", "119"
                    DMM "TRIG", "SGL"
                    
                    If FixedRdg >= 29.8 And FixedRdg <= 32 Then
                    ActiveCell.Value = "Pass"
                    Else
                    ActiveCell.Value = "Fail"
                    End If
                    Selection.OffSet(1, 0).Select
                    End If
                    TestSect = TestSectBak
                    PrevTestSect = TestSect
                    
                
'250 ohm HART Resistor

            Case "$F$89", "$G$89"
            
                TestSections 16
                
                
                '-------------------Begin Shows the MainHookup Userform for hookups and pictures---------------------
                If PrevTestSect <> TestSect Then UForms "MainForm"
                If TerminateClicked Then TerminateClicked = False: Exit Sub
                '-------------------End Shows the MainHookup Userform for hookups and pictures-----------------------
                    If DMMGPIB = "" Then
                    TestSections 9000
                    SectTitleString = "B89"
                    CommentsString = "Is the Output voltage from source jacks between 29.8 to 32V?"
                    If PrevTestSect <> TestSect Then UForms "MainForm_Basic_Comment"
                    If TerminateClicked Then TerminateClicked = False: Exit Sub
                    Else
                    AutoSelect = False
                    Multiplier = 1
                    DMMClearStatus "Reset"
                    DMM "END", "ALWAYS"
                    DMM "Func", "DCV"
                    DMMReset = 1
                    DMM "Range", "119"
                    DMM "TRIG", "SGL"
                    If FixedRdg >= 29.8 And FixedRdg <= 32 Then
                    ActiveCell.Value = "Pass"
                    Else
                    ActiveCell.Value = "Fail"
                    End If
                    Selection.OffSet(1, 0).Select
                    End If
                    TestSect = TestSectBak
                    PrevTestSect = TestSect
                
'1 kohm Shunt Resistor

            Case "$F$90", "$G$90"
            
                TestSections 17
                '-------------------Begin Shows the MainHookup Userform for hookups and pictures---------------------
                If PrevTestSect <> TestSect Then UForms "MainForm"
                If TerminateClicked Then TerminateClicked = False: Exit Sub
                '-------------------End Shows the MainHookup Userform for hookups and pictures-----------------------
                    If DMMGPIB = "" Then
                    TestSections 9000
                    SectTitleString = "B90"
                    CommentsString = "Is the Output voltage from source jacks between 23.8 to 32V?"
                    If PrevTestSect <> TestSect Then UForms "MainForm_Basic_Comment"
                    If TerminateClicked Then TerminateClicked = False: Exit Sub
                    Else
                    AutoSelect = False
                    Multiplier = 1
                    DMMClearStatus "Reset"
                    DMM "END", "ALWAYS"
                    DMM "Func", "DCV"
                    DMMReset = 1
                    DMM "Range", "119"
                    DMM "TRIG", "SGL"
                    If FixedRdg >= 23.8 And FixedRdg <= 32 Then
                    ActiveCell.Value = "Pass"
                    Else
                    ActiveCell.Value = "Fail"
                    End If
                    Selection.OffSet(1, 0).Select
                    End If
                    TestSect = TestSectBak
                    PrevTestSect = TestSect
                
'Current Source

            Case "$F$91", "$G$91"
            
                TestSections 18
                '-------------------Begin Shows the MainHookup Userform for hookups and pictures---------------------
                If PrevTestSect <> TestSect Then UForms "MainForm"
                If TerminateClicked Then TerminateClicked = False: Exit Sub
                '-------------------End Shows the MainHookup Userform for hookups and pictures-----------------------
                
                    If DMMGPIB = "" Then
                    TestSections 9000
                    SectTitleString = "B91"
                    CommentsString = "Is the Output current from source jacks between 24.5 to 35 mA?"
                    If PrevTestSect <> TestSect Then UForms "MainForm_Basic_Comment"
                    If TerminateClicked Then TerminateClicked = False: Exit Sub
                    Else
                    AutoSelect = False
                    Multiplier = 1000
                    DMMClearStatus "Reset"
                    DMM "END", "ALWAYS"
                    DMM "Func", "DCI"
                    DMMReset = 1
                    DMM "Range", "100E-3"
                    DMM "TRIG", "SGL"
                    If FixedRdg >= 24.5 And FixedRdg <= 35 Then
                    ActiveCell.Value = "Pass"
                    Else
                    ActiveCell.Value = "Fail"
                    End If
                    Selection.OffSet(1, 0).Select
                    End If
                    TestSect = TestSectBak
                    PrevTestSect = TestSect
                    

                
                
                
                
                
                '-------------------New Code Ends Here----------------
                


 

            Case "$F$92", "$G$92"
                '-------------------End Here on last As Found/Left Gray Cells---------------------
                CalibClearStatus "Standby"
                MsgBox "Verification Complete! Remove All Connections!"
                ActiveSheet.Range("I9").Select
                
            Case Else
                '-------------------Clicking anywhere else---------------------
                If ToggleStates("CodeButton") = "Off" Then
                
                ElseIf ToggleStates("CodeButton") = "Operating" Then
                    ButtonState PanelForm, "CodeButton", "Standby"
                    CalibClearStatus "Standby"
                    CalibClearStatus "Close"
                    DMMClearStatus "Close"
                ElseIf ToggleStates("CodeButton") = "Standby" Then
                
                End If
                
                DMMReset = 0
                TestSect = 0
                TestForm = 0
                PrevTestSect = 0
                volt = 0
                Sheet2.Range("I9").Value = ""
                Sheet2.Range("AA1").Value = "Standby"
                HVImageShow volt, "V"
                
                
        End Select
End If
    
'---------------------------------------------------------------------------------
'---------------------------------Edit Above--------------------------------------
'---------------------------------------------------------------------------------
    
    Exit Sub
    
ErrorHandler:
    Application.EnableEvents = True
    MsgBox "Error: " & err.Description
    Call ReportError("YourMacroName", err.Number, err.Description, Erl)
    
End Sub
Sub ReportError(procName As String, ErrNum As Long, ErrDesc As String, ErrLine As Long)
    MsgBox "Error in " & procName & vbCrLf & _
           "Line: " & ErrLine & vbCrLf & _
           "Error " & ErrNum & ": " & ErrDesc, vbCritical
End Sub

Sub JMTestHRSSetup()
    'This is the Main Setup for the DMM (3458A)
    'MsgBox DMMReset
    If DMMReset = 1 Then Exit Sub
    
    DMMClearStatus "Reset"
    DMM "END", "ALWAYS"
    DMM "NPLC", "10"
    DMM "NRDGS", "30"
    DMM "Func", "DCI"
    DMMReset = 1
    
End Sub

