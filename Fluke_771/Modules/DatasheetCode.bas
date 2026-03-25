Attribute VB_Name = "DatasheetCode"
' Fluke 771 Milliamp Process Clamp Meter - DatasheetCode
' DC mA Clamp Meter - Calibrator sources current through loop, clamp measures

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
    '-------------------End Standard Reset Check----------------------------------

'---------------------------------------------------------------------------------
'---------------------------------Edit Below--------------------------------------
'---------------------------------------------------------------------------------

If PanelForm.CodeButton.Caption = "Off" Then

Else
'Fluke 771 DC mA Clamp Meter
'Calibrator sources DC current, clamp meter measures around the loop
'Command format: Calibrator "Source", "DCI", value, "mA", 0, "Hz", "", 0, 0, ""

        '-------------------Begin All Test Points on Datasheet---------------------
        If Target Is Nothing Then Exit Sub
        Select Case Target.Address

'Operational Checks (Pass/Fail only - no calibrator commands)
            Case "$F$14", "$G$14"
                'Backlight Test
                TestSections 1000
                Selection.Offset(1, 0).Select

            Case "$F$15", "$G$15"
                'Display Test
                TestSections 1000
                Selection.Offset(1, 0).Select

            Case "$F$16", "$G$16"
                'Keypad Test
                TestSections 1000
                Selection.Offset(1, 0).Select

            Case "$F$17", "$G$17"
                'Spotlight Test
                TestSections 1000
                Selection.Offset(1, 0).Select

'DC Current Tests - 20.99 mA Range
            Case "$F$20", "$G$20"
                '+4 mA
                TestSections 6000
                '-------------------Begin Shows the MainHookup Userform for hookups and pictures---------------------
                If PrevTestSect <> TestSect Then UForms "MainForm"
                If TerminateClicked Then TerminateClicked = False: Exit Sub
                '-------------------End Shows the MainHookup Userform for hookups and pictures-----------------------

                Calibrator "Source", "DCI", 4, "mA", 0, "Hz", "", 0, 0, ""
                PrevTestSect = TestSect

            Case "$F$21", "$G$21"
                '-4 mA
                TestSections 6000
                '-------------------Begin Shows the MainHookup Userform for hookups and pictures---------------------
                If PrevTestSect <> TestSect Then UForms "MainForm"
                If TerminateClicked Then TerminateClicked = False: Exit Sub
                '-------------------End Shows the MainHookup Userform for hookups and pictures-----------------------

                Calibrator "Source", "DCI", -4, "mA", 0, "Hz", "", 0, 0, ""
                PrevTestSect = TestSect

            Case "$F$22", "$G$22"
                '+12 mA
                TestSections 6000
                '-------------------Begin Shows the MainHookup Userform for hookups and pictures---------------------
                If PrevTestSect <> TestSect Then UForms "MainForm"
                If TerminateClicked Then TerminateClicked = False: Exit Sub
                '-------------------End Shows the MainHookup Userform for hookups and pictures-----------------------

                Calibrator "Source", "DCI", 12, "mA", 0, "Hz", "", 0, 0, ""
                PrevTestSect = TestSect

            Case "$F$23", "$G$23"
                '-12 mA
                TestSections 6000
                '-------------------Begin Shows the MainHookup Userform for hookups and pictures---------------------
                If PrevTestSect <> TestSect Then UForms "MainForm"
                If TerminateClicked Then TerminateClicked = False: Exit Sub
                '-------------------End Shows the MainHookup Userform for hookups and pictures-----------------------

                Calibrator "Source", "DCI", -12, "mA", 0, "Hz", "", 0, 0, ""
                PrevTestSect = TestSect

            Case "$F$24", "$G$24"
                '+20 mA
                TestSections 6000
                '-------------------Begin Shows the MainHookup Userform for hookups and pictures---------------------
                If PrevTestSect <> TestSect Then UForms "MainForm"
                If TerminateClicked Then TerminateClicked = False: Exit Sub
                '-------------------End Shows the MainHookup Userform for hookups and pictures-----------------------

                Calibrator "Source", "DCI", 20, "mA", 0, "Hz", "", 0, 0, ""
                PrevTestSect = TestSect

            Case "$F$25", "$G$25"
                '-20 mA
                TestSections 6000
                '-------------------Begin Shows the MainHookup Userform for hookups and pictures---------------------
                If PrevTestSect <> TestSect Then UForms "MainForm"
                If TerminateClicked Then TerminateClicked = False: Exit Sub
                '-------------------End Shows the MainHookup Userform for hookups and pictures-----------------------

                Calibrator "Source", "DCI", -20, "mA", 0, "Hz", "", 0, 0, ""
                PrevTestSect = TestSect

'DC Current Tests - 99.9 mA Range
            Case "$F$27", "$G$27"
                '+100 mA
                TestSections 6000
                '-------------------Begin Shows the MainHookup Userform for hookups and pictures---------------------
                If PrevTestSect <> TestSect Then UForms "MainForm"
                If TerminateClicked Then TerminateClicked = False: Exit Sub
                '-------------------End Shows the MainHookup Userform for hookups and pictures-----------------------

                Calibrator "Source", "DCI", 100, "mA", 0, "Hz", "", 0, 0, ""
                PrevTestSect = TestSect

            Case "$F$28", "$G$28"
                '-100 mA
                TestSections 6000
                '-------------------Begin Shows the MainHookup Userform for hookups and pictures---------------------
                If PrevTestSect <> TestSect Then UForms "MainForm"
                If TerminateClicked Then TerminateClicked = False: Exit Sub
                '-------------------End Shows the MainHookup Userform for hookups and pictures-----------------------

                Calibrator "Source", "DCI", -100, "mA", 0, "Hz", "", 0, 0, ""
                PrevTestSect = TestSect

            Case "$F$29", "$G$29"
                '-------------------End Here on last As Found/Left Gray Cells---------------------
                CalibClearStatus "Standby"
                MsgBox "Verification Complete! Remove All Connections!"
                ActiveSheet.Range("I9").Select

            Case Else
                '-------------------Clicking anywhere else---------------------
                If PanelForm.CodeButton.Caption = "Off" Then

                ElseIf PanelForm.CodeButton.Caption = "Operating" Then
                    ButtonState PanelForm, "CodeButton", "Standby"
                    CalibClearStatus "Standby"
                    CalibClearStatus "Close"
                    DMMClearStatus "Close"
                ElseIf PanelForm.CodeButton.Caption = "Standby" Then

                End If

                DMMReset = 0
                TestSect = 0
                TestForm = 0
                PrevTestSect = 0
                volt = 0
                HVImageShow volt, "V"


        End Select
End If

'---------------------------------------------------------------------------------
'---------------------------------Edit Above--------------------------------------
'---------------------------------------------------------------------------------

    Exit Sub

ErrorHandler:
    Application.EnableEvents = True
    MsgBox "Error: " & Err.Description
    Call ReportError("HandleSelectionChange", Err.Number, Err.Description, Erl)

End Sub

Sub ReportError(procName As String, ErrNum As Long, ErrDesc As String, ErrLine As Long)
    MsgBox "Error in " & procName & vbCrLf & _
           "Line: " & ErrLine & vbCrLf & _
           "Error " & ErrNum & ": " & ErrDesc, vbCritical
End Sub
