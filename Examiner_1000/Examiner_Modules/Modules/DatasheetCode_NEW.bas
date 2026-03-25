Attribute VB_Name = "DatasheetCode"

Public Sub HandleSelectionChange(ByVal Target As Excel.Range)
    On Error GoTo ErrorHandler
    Dim i           As Long
    Dim startRow    As Long
    Dim endRow      As Long
    Dim currentRow  As Long
    Dim btn         As Shape
    Dim btnExists   As Boolean
    Dim shp         As Shape
    Dim volt        As Double

    ' Get calibrator model info (if using Info sheet like 789)
    ' CalibratorModel = wsInfo.Range("M9").Value
    ' DMMModel = wsInfo.Range("P9").Value
    ' CounterModel = wsInfo.Range("M16").Value

    SetupWS

    ' Check if CommToggle button exists
    btnExists = False
    For Each shp In Sheet2.Shapes
        If shp.Name = "CommToggle" Then
            btnExists = True
            Set btn = shp
            Exit For
        End If
    Next shp

    If Not btnExists Then
        MsgBox "CommToggle button not found on Sheet2.", vbExclamation
        Exit Sub
    End If

    ' Initialize calibrator on first use
    If Sheet2.Range("AA1").Value = "" Then
        Comm True, False, False
        Sheet2.Range("AA1").Value = "Standby"
    End If

'---------------------------------------------------------------------------------
'---------------------------------Edit Below--------------------------------------
'---------------------------------------------------------------------------------

If Sheet2.Range("AA1").Value = "Standby" Or Sheet2.Range("AA1").Value = "Operating" Then

    '-------------------Begin All Test Points on Datasheet---------------------
    If Target Is Nothing Then Exit Sub
    Select Case Target.Address

        '========================================================================
        ' TestSect 1 (SameTest=1): AC Voltage with varying frequencies
        '========================================================================

        Case "$G$20", "$H$20"
            TestSect = 1
            If PrevSameTest <> SameTest Then
                SameTest = 1
                If PrevSameTest <> SameTest Then
                    MainHookup.Show
                    If TerminateClicked Then TerminateClicked = False: Exit Sub
                End If
            End If
            With btn
                .TextFrame.Characters.Text = "Operating"
                With .TextFrame.Characters.Font
                    .Size = 20
                    .Bold = True
                    .Color = RGB(255, 0, 0)
                End With
            End With
            Sheet2.Range("AA1").Value = "Operating"
            HVImageShow
            Comm True, False, False
            instrument.WriteString "*cls"
            instrument.WriteString "OUT " & 0.1 & " " & "V" & ", " & 10 & " " & "Hz" & "; OPER"
            PrevTestSect = TestSect
            PrevSameTest = SameTest

        Case "$G$21", "$H$21"
            TestSect = 1
            SameTest = 1
            With btn
                .TextFrame.Characters.Text = "Operating"
                With .TextFrame.Characters.Font
                    .Size = 20
                    .Bold = True
                    .Color = RGB(255, 0, 0)
                End With
            End With
            Sheet2.Range("AA1").Value = "Operating"
            HVImageShow
            Comm True, False, False
            instrument.WriteString "*cls"
            instrument.WriteString "OUT " & 0.1 & " " & "V" & ", " & 50 & " " & "Hz" & "; OPER"
            PrevTestSect = TestSect
            PrevSameTest = SameTest

        Case "$G$22", "$H$22"
            TestSect = 1
            SameTest = 1
            With btn
                .TextFrame.Characters.Text = "Operating"
                With .TextFrame.Characters.Font
                    .Size = 20
                    .Bold = True
                    .Color = RGB(255, 0, 0)
                End With
            End With
            Sheet2.Range("AA1").Value = "Operating"
            HVImageShow
            Comm True, False, False
            instrument.WriteString "*cls"
            instrument.WriteString "OUT " & 0.1 & " " & "V" & ", " & 100 & " " & "Hz" & "; OPER"
            PrevTestSect = TestSect
            PrevSameTest = SameTest

        Case "$G$23", "$H$23"
            TestSect = 1
            SameTest = 1
            With btn
                .TextFrame.Characters.Text = "Operating"
                With .TextFrame.Characters.Font
                    .Size = 20
                    .Bold = True
                    .Color = RGB(255, 0, 0)
                End With
            End With
            Sheet2.Range("AA1").Value = "Operating"
            HVImageShow
            Comm True, False, False
            instrument.WriteString "*cls"
            instrument.WriteString "OUT " & 0.1 & " " & "V" & ", " & 500 & " " & "Hz" & "; OPER"
            PrevTestSect = TestSect
            PrevSameTest = SameTest

        Case "$G$24", "$H$24"
            TestSect = 1
            SameTest = 1
            With btn
                .TextFrame.Characters.Text = "Operating"
                With .TextFrame.Characters.Font
                    .Size = 20
                    .Bold = True
                    .Color = RGB(255, 0, 0)
                End With
            End With
            Sheet2.Range("AA1").Value = "Operating"
            HVImageShow
            Comm True, False, False
            instrument.WriteString "*cls"
            instrument.WriteString "OUT " & 0.1 & " " & "V" & ", " & 1 & " " & "kHz" & "; OPER"
            PrevTestSect = TestSect
            PrevSameTest = SameTest

        Case "$G$25", "$H$25"
            TestSect = 1
            SameTest = 1
            With btn
                .TextFrame.Characters.Text = "Operating"
                With .TextFrame.Characters.Font
                    .Size = 20
                    .Bold = True
                    .Color = RGB(255, 0, 0)
                End With
            End With
            Sheet2.Range("AA1").Value = "Operating"
            HVImageShow
            Comm True, False, False
            instrument.WriteString "*cls"
            instrument.WriteString "OUT " & 0.1 & " " & "V" & ", " & 5 & " " & "kHz" & "; OPER"
            PrevTestSect = TestSect
            PrevSameTest = SameTest

        Case "$G$26", "$H$26"
            TestSect = 1
            SameTest = 1
            With btn
                .TextFrame.Characters.Text = "Operating"
                With .TextFrame.Characters.Font
                    .Size = 20
                    .Bold = True
                    .Color = RGB(255, 0, 0)
                End With
            End With
            Sheet2.Range("AA1").Value = "Operating"
            HVImageShow
            Comm True, False, False
            instrument.WriteString "*cls"
            instrument.WriteString "OUT " & 0.1 & " " & "V" & ", " & 10 & " " & "kHz" & "; OPER"
            PrevTestSect = TestSect
            PrevSameTest = SameTest

        '-------------------Skip Cells---------------------
        Case "$G$27", "$H$27"
            On Error Resume Next
            ActiveCell.Offset(1, 0).Select
            If Err.Number <> 0 Then
                MsgBox "Error selecting next cell: " & Err.Description
                Err.Clear
            End If
            On Error GoTo 0

        '-------------------Standby Skip Cells---------------------
        Case "$G$28", "$H$28"
            CommToggle "Standby"
            Comm False, True, False
            Cls
            On Error Resume Next
            ActiveCell.Offset(1, 0).Select
            If Err.Number <> 0 Then
                MsgBox "Error selecting next cell: " & Err.Description
                Err.Clear
            End If
            On Error GoTo 0

        '========================================================================
        ' TestSect 2 (SameTest=1): AC Voltage at 100 Hz with varying amplitudes
        '========================================================================

        Case "$G$29", "$H$29"
            TestSect = 2
            If PrevSameTest <> SameTest Then
                SameTest = 1
                If PrevSameTest <> SameTest Then
                    MainHookup.Show
                    If TerminateClicked Then TerminateClicked = False: Exit Sub
                End If
            End If
            With btn
                .TextFrame.Characters.Text = "Operating"
                With .TextFrame.Characters.Font
                    .Size = 20
                    .Bold = True
                    .Color = RGB(255, 0, 0)
                End With
            End With
            Sheet2.Range("AA1").Value = "Operating"
            HVImageShow
            Comm True, False, False
            instrument.WriteString "*cls"
            instrument.WriteString "OUT " & 0.005 & " " & "V" & ", " & 100 & " " & "Hz" & "; OPER"
            PrevTestSect = TestSect
            PrevSameTest = SameTest

        Case "$G$30", "$H$30"
            TestSect = 2
            SameTest = 1
            With btn
                .TextFrame.Characters.Text = "Operating"
                With .TextFrame.Characters.Font
                    .Size = 20
                    .Bold = True
                    .Color = RGB(255, 0, 0)
                End With
            End With
            Sheet2.Range("AA1").Value = "Operating"
            HVImageShow
            Comm True, False, False
            instrument.WriteString "*cls"
            instrument.WriteString "OUT " & 0.01 & " " & "V" & ", " & 100 & " " & "Hz" & "; OPER"
            PrevTestSect = TestSect
            PrevSameTest = SameTest

        Case "$G$31", "$H$31"
            TestSect = 2
            SameTest = 1
            With btn
                .TextFrame.Characters.Text = "Operating"
                With .TextFrame.Characters.Font
                    .Size = 20
                    .Bold = True
                    .Color = RGB(255, 0, 0)
                End With
            End With
            Sheet2.Range("AA1").Value = "Operating"
            HVImageShow
            Comm True, False, False
            instrument.WriteString "*cls"
            instrument.WriteString "OUT " & 0.05 & " " & "V" & ", " & 100 & " " & "Hz" & "; OPER"
            PrevTestSect = TestSect
            PrevSameTest = SameTest

        Case "$G$32", "$H$32"
            TestSect = 2
            SameTest = 1
            With btn
                .TextFrame.Characters.Text = "Operating"
                With .TextFrame.Characters.Font
                    .Size = 20
                    .Bold = True
                    .Color = RGB(255, 0, 0)
                End With
            End With
            Sheet2.Range("AA1").Value = "Operating"
            HVImageShow
            Comm True, False, False
            instrument.WriteString "*cls"
            instrument.WriteString "OUT " & 0.1 & " " & "V" & ", " & 100 & " " & "Hz" & "; OPER"
            PrevTestSect = TestSect
            PrevSameTest = SameTest

        Case "$G$33", "$H$33"
            TestSect = 2
            SameTest = 1
            With btn
                .TextFrame.Characters.Text = "Operating"
                With .TextFrame.Characters.Font
                    .Size = 20
                    .Bold = True
                    .Color = RGB(255, 0, 0)
                End With
            End With
            Sheet2.Range("AA1").Value = "Operating"
            HVImageShow
            Comm True, False, False
            instrument.WriteString "*cls"
            instrument.WriteString "OUT " & 0.2 & " " & "V" & ", " & 100 & " " & "Hz" & "; OPER"
            PrevTestSect = TestSect
            PrevSameTest = SameTest

        Case "$G$34", "$H$34"
            TestSect = 2
            SameTest = 1
            With btn
                .TextFrame.Characters.Text = "Operating"
                With .TextFrame.Characters.Font
                    .Size = 20
                    .Bold = True
                    .Color = RGB(255, 0, 0)
                End With
            End With
            Sheet2.Range("AA1").Value = "Operating"
            HVImageShow
            Comm True, False, False
            instrument.WriteString "*cls"
            instrument.WriteString "OUT " & 0.5 & " " & "V" & ", " & 100 & " " & "Hz" & "; OPER"
            PrevTestSect = TestSect
            PrevSameTest = SameTest

        Case "$G$35", "$H$35"
            TestSect = 2
            SameTest = 1
            With btn
                .TextFrame.Characters.Text = "Operating"
                With .TextFrame.Characters.Font
                    .Size = 20
                    .Bold = True
                    .Color = RGB(255, 0, 0)
                End With
            End With
            Sheet2.Range("AA1").Value = "Operating"
            HVImageShow
            Comm True, False, False
            instrument.WriteString "*cls"
            instrument.WriteString "OUT " & 1 & " " & "V" & ", " & 100 & " " & "Hz" & "; OPER"
            PrevTestSect = TestSect
            PrevSameTest = SameTest

        Case "$G$36", "$H$36"
            TestSect = 2
            SameTest = 1
            With btn
                .TextFrame.Characters.Text = "Operating"
                With .TextFrame.Characters.Font
                    .Size = 20
                    .Bold = True
                    .Color = RGB(255, 0, 0)
                End With
            End With
            Sheet2.Range("AA1").Value = "Operating"
            HVImageShow
            Comm True, False, False
            instrument.WriteString "*cls"
            instrument.WriteString "OUT " & 1.5 & " " & "V" & ", " & 100 & " " & "Hz" & "; OPER"
            PrevTestSect = TestSect
            PrevSameTest = SameTest

        Case "$G$37", "$H$37"
            TestSect = 2
            SameTest = 1
            With btn
                .TextFrame.Characters.Text = "Operating"
                With .TextFrame.Characters.Font
                    .Size = 20
                    .Bold = True
                    .Color = RGB(255, 0, 0)
                End With
            End With
            Sheet2.Range("AA1").Value = "Operating"
            HVImageShow
            Comm True, False, False
            instrument.WriteString "*cls"
            instrument.WriteString "OUT " & 1.95 & " " & "V" & ", " & 100 & " " & "Hz" & "; OPER"
            PrevTestSect = TestSect
            PrevSameTest = SameTest

        '-------------------Skip Cells---------------------
        Case "$G$38", "$H$38"
            On Error Resume Next
            ActiveCell.Offset(1, 0).Select
            If Err.Number <> 0 Then
                MsgBox "Error selecting next cell: " & Err.Description
                Err.Clear
            End If
            On Error GoTo 0

        '-------------------Standby Skip Cells---------------------
        Case "$G$39", "$H$39"
            CommToggle "Standby"
            Comm False, True, False
            Cls
            On Error Resume Next
            ActiveCell.Offset(1, 0).Select
            If Err.Number <> 0 Then
                MsgBox "Error selecting next cell: " & Err.Description
                Err.Clear
            End If
            On Error GoTo 0

        '========================================================================
        ' TestSect 3 (SameTest=2): AC Voltage with varying amplitudes and frequencies
        '========================================================================

        Case "$G$40", "$H$40"
            TestSect = 3
            If PrevSameTest <> SameTest Then
                SameTest = 2
                If PrevSameTest <> SameTest Then
                    MainHookup.Show
                    If TerminateClicked Then TerminateClicked = False: Exit Sub
                End If
            End If
            With btn
                .TextFrame.Characters.Text = "Operating"
                With .TextFrame.Characters.Font
                    .Size = 20
                    .Bold = True
                    .Color = RGB(255, 0, 0)
                End With
            End With
            Sheet2.Range("AA1").Value = "Operating"
            HVImageShow
            Comm True, False, False
            instrument.WriteString "*cls"
            instrument.WriteString "OUT " & 0.1 & " " & "V" & ", " & 10 & " " & "Hz" & "; OPER"
            PrevTestSect = TestSect
            PrevSameTest = SameTest

        Case "$G$41", "$H$41"
            TestSect = 3
            SameTest = 2
            With btn
                .TextFrame.Characters.Text = "Operating"
                With .TextFrame.Characters.Font
                    .Size = 20
                    .Bold = True
                    .Color = RGB(255, 0, 0)
                End With
            End With
            Sheet2.Range("AA1").Value = "Operating"
            HVImageShow
            Comm True, False, False
            instrument.WriteString "*cls"
            instrument.WriteString "OUT " & 0.1 & " " & "V" & ", " & 15 & " " & "Hz" & "; OPER"
            PrevTestSect = TestSect
            PrevSameTest = SameTest

        Case "$G$42", "$H$42"
            TestSect = 3
            SameTest = 2
            With btn
                .TextFrame.Characters.Text = "Operating"
                With .TextFrame.Characters.Font
                    .Size = 20
                    .Bold = True
                    .Color = RGB(255, 0, 0)
                End With
            End With
            Sheet2.Range("AA1").Value = "Operating"
            HVImageShow
            Comm True, False, False
            instrument.WriteString "*cls"
            instrument.WriteString "OUT " & 0.1 & " " & "V" & ", " & 20 & " " & "Hz" & "; OPER"
            PrevTestSect = TestSect
            PrevSameTest = SameTest

        Case "$G$43", "$H$43"
            TestSect = 3
            SameTest = 2
            With btn
                .TextFrame.Characters.Text = "Operating"
                With .TextFrame.Characters.Font
                    .Size = 20
                    .Bold = True
                    .Color = RGB(255, 0, 0)
                End With
            End With
            Sheet2.Range("AA1").Value = "Operating"
            HVImageShow
            Comm True, False, False
            instrument.WriteString "*cls"
            instrument.WriteString "OUT " & 0.2 & " " & "V" & ", " & 25 & " " & "Hz" & "; OPER"
            PrevTestSect = TestSect
            PrevSameTest = SameTest

        Case "$G$44", "$H$44"
            TestSect = 3
            SameTest = 2
            With btn
                .TextFrame.Characters.Text = "Operating"
                With .TextFrame.Characters.Font
                    .Size = 20
                    .Bold = True
                    .Color = RGB(255, 0, 0)
                End With
            End With
            Sheet2.Range("AA1").Value = "Operating"
            HVImageShow
            Comm True, False, False
            instrument.WriteString "*cls"
            instrument.WriteString "OUT " & 0.2 & " " & "V" & ", " & 30 & " " & "Hz" & "; OPER"
            PrevTestSect = TestSect
            PrevSameTest = SameTest

        Case "$G$45", "$H$45"
            TestSect = 3
            SameTest = 2
            With btn
                .TextFrame.Characters.Text = "Operating"
                With .TextFrame.Characters.Font
                    .Size = 20
                    .Bold = True
                    .Color = RGB(255, 0, 0)
                End With
            End With
            Sheet2.Range("AA1").Value = "Operating"
            HVImageShow
            Comm True, False, False
            instrument.WriteString "*cls"
            instrument.WriteString "OUT " & 0.2 & " " & "V" & ", " & 50 & " " & "Hz" & "; OPER"
            PrevTestSect = TestSect
            PrevSameTest = SameTest

        Case "$G$46", "$H$46"
            TestSect = 3
            SameTest = 2
            With btn
                .TextFrame.Characters.Text = "Operating"
                With .TextFrame.Characters.Font
                    .Size = 20
                    .Bold = True
                    .Color = RGB(255, 0, 0)
                End With
            End With
            Sheet2.Range("AA1").Value = "Operating"
            HVImageShow
            Comm True, False, False
            instrument.WriteString "*cls"
            instrument.WriteString "OUT " & 0.3 & " " & "V" & ", " & 100 & " " & "Hz" & "; OPER"
            PrevTestSect = TestSect
            PrevSameTest = SameTest

        Case "$G$47", "$H$47"
            TestSect = 3
            SameTest = 2
            With btn
                .TextFrame.Characters.Text = "Operating"
                With .TextFrame.Characters.Font
                    .Size = 20
                    .Bold = True
                    .Color = RGB(255, 0, 0)
                End With
            End With
            Sheet2.Range("AA1").Value = "Operating"
            HVImageShow
            Comm True, False, False
            instrument.WriteString "*cls"
            instrument.WriteString "OUT " & 0.6 & " " & "V" & ", " & 500 & " " & "Hz" & "; OPER"
            PrevTestSect = TestSect
            PrevSameTest = SameTest

        Case "$G$48", "$H$48"
            TestSect = 3
            SameTest = 2
            With btn
                .TextFrame.Characters.Text = "Operating"
                With .TextFrame.Characters.Font
                    .Size = 20
                    .Bold = True
                    .Color = RGB(255, 0, 0)
                End With
            End With
            Sheet2.Range("AA1").Value = "Operating"
            HVImageShow
            Comm True, False, False
            instrument.WriteString "*cls"
            instrument.WriteString "OUT " & 1.2 & " " & "V" & ", " & 1 & " " & "kHz" & "; OPER"
            PrevTestSect = TestSect
            PrevSameTest = SameTest

        Case "$G$49", "$H$49"
            TestSect = 3
            SameTest = 2
            With btn
                .TextFrame.Characters.Text = "Operating"
                With .TextFrame.Characters.Font
                    .Size = 20
                    .Bold = True
                    .Color = RGB(255, 0, 0)
                End With
            End With
            Sheet2.Range("AA1").Value = "Operating"
            HVImageShow
            Comm True, False, False
            instrument.WriteString "*cls"
            instrument.WriteString "OUT " & 1.5 & " " & "V" & ", " & 2 & " " & "kHz" & "; OPER"
            PrevTestSect = TestSect
            PrevSameTest = SameTest

        Case "$G$50", "$H$50"
            TestSect = 3
            SameTest = 2
            With btn
                .TextFrame.Characters.Text = "Operating"
                With .TextFrame.Characters.Font
                    .Size = 20
                    .Bold = True
                    .Color = RGB(255, 0, 0)
                End With
            End With
            Sheet2.Range("AA1").Value = "Operating"
            HVImageShow
            Comm True, False, False
            instrument.WriteString "*cls"
            instrument.WriteString "OUT " & 1.9 & " " & "V" & ", " & 5 & " " & "kHz" & "; OPER"
            PrevTestSect = TestSect
            PrevSameTest = SameTest

        Case "$G$51", "$H$51"
            TestSect = 3
            SameTest = 2
            With btn
                .TextFrame.Characters.Text = "Operating"
                With .TextFrame.Characters.Font
                    .Size = 20
                    .Bold = True
                    .Color = RGB(255, 0, 0)
                End With
            End With
            Sheet2.Range("AA1").Value = "Operating"
            HVImageShow
            Comm True, False, False
            instrument.WriteString "*cls"
            instrument.WriteString "OUT " & 1.9 & " " & "V" & ", " & 10 & " " & "kHz" & "; OPER"
            PrevTestSect = TestSect
            PrevSameTest = SameTest

        '-------------------Skip Cells---------------------
        Case "$G$52", "$H$52"
            On Error Resume Next
            ActiveCell.Offset(1, 0).Select
            If Err.Number <> 0 Then
                MsgBox "Error selecting next cell: " & Err.Description
                Err.Clear
            End If
            On Error GoTo 0

        '-------------------Standby Skip Cells---------------------
        Case "$G$53", "$H$53"
            CommToggle "Standby"
            Comm False, True, False
            Cls
            On Error Resume Next
            ActiveCell.Offset(1, 0).Select
            If Err.Number <> 0 Then
                MsgBox "Error selecting next cell: " & Err.Description
                Err.Clear
            End If
            On Error GoTo 0

        '========================================================================
        ' TestSect 4 (SameTest=3): Single test point
        '========================================================================

        Case "$G$54", "$H$54"
            TestSect = 4
            If PrevSameTest <> SameTest Then
                SameTest = 3
                If PrevSameTest <> SameTest Then
                    MainHookup.Show
                    If TerminateClicked Then TerminateClicked = False: Exit Sub
                End If
            End If
            With btn
                .TextFrame.Characters.Text = "Operating"
                With .TextFrame.Characters.Font
                    .Size = 20
                    .Bold = True
                    .Color = RGB(255, 0, 0)
                End With
            End With
            Sheet2.Range("AA1").Value = "Operating"
            HVImageShow
            Comm True, False, False
            instrument.WriteString "*cls"
            instrument.WriteString "OUT " & 0.1 & " " & "V" & ", " & 500 & " " & "Hz" & "; OPER"
            PrevTestSect = TestSect
            PrevSameTest = SameTest

        '========================================================================
        ' TestSect 5 (SameTest=3): Single test point
        '========================================================================

        Case "$G$55", "$H$55"
            TestSect = 5
            SameTest = 3
            With btn
                .TextFrame.Characters.Text = "Operating"
                With .TextFrame.Characters.Font
                    .Size = 20
                    .Bold = True
                    .Color = RGB(255, 0, 0)
                End With
            End With
            Sheet2.Range("AA1").Value = "Operating"
            HVImageShow
            Comm True, False, False
            instrument.WriteString "*cls"
            instrument.WriteString "OUT " & 0.1 & " " & "V" & ", " & 5 & " " & "kHz" & "; OPER"
            PrevTestSect = TestSect
            PrevSameTest = SameTest

        '-------------------Skip Cells---------------------
        Case "$G$56", "$H$56"
            On Error Resume Next
            ActiveCell.Offset(1, 0).Select
            If Err.Number <> 0 Then
                MsgBox "Error selecting next cell: " & Err.Description
                Err.Clear
            End If
            On Error GoTo 0

        '-------------------Standby Skip Cells---------------------
        Case "$G$57", "$H$57"
            CommToggle "Standby"
            Comm False, True, False
            Cls
            On Error Resume Next
            ActiveCell.Offset(1, 0).Select
            If Err.Number <> 0 Then
                MsgBox "Error selecting next cell: " & Err.Description
                Err.Clear
            End If
            On Error GoTo 0

        '========================================================================
        ' TestSect 6 (SameTest=4): Single test point - 440 Hz
        '========================================================================

        Case "$G$58", "$H$58"
            TestSect = 6
            If PrevSameTest <> SameTest Then
                SameTest = 4
                If PrevSameTest <> SameTest Then
                    MainHookup.Show
                    If TerminateClicked Then TerminateClicked = False: Exit Sub
                End If
            End If
            With btn
                .TextFrame.Characters.Text = "Operating"
                With .TextFrame.Characters.Font
                    .Size = 20
                    .Bold = True
                    .Color = RGB(255, 0, 0)
                End With
            End With
            Sheet2.Range("AA1").Value = "Operating"
            HVImageShow
            Comm True, False, False
            instrument.WriteString "*cls"
            instrument.WriteString "OUT " & 0.1 & " " & "V" & ", " & 440 & " " & "Hz" & "; OPER"
            PrevTestSect = TestSect
            PrevSameTest = SameTest

        '========================================================================
        ' TestSect 7: Pass/Fail userform - Volume control check
        '========================================================================

        Case "$G$59", "$H$59"
            TestSect = 7
            ' Show Pass/Fail userform for volume control check
            ' Tech listens to verify volume control works while DUT is operating
            ' (MainHookup should have Pass/Fail buttons for this)
            MainHookup.Show
            If TerminateClicked Then TerminateClicked = False: Exit Sub
            PrevTestSect = TestSect

        '-------------------Standby Skip Cells---------------------
        Case "$G$60", "$H$60"
            CommToggle "Standby"
            Comm False, True, False
            Cls
            On Error Resume Next
            ActiveCell.Offset(1, 0).Select
            If Err.Number <> 0 Then
                MsgBox "Error selecting next cell: " & Err.Description
                Err.Clear
            End If
            On Error GoTo 0

        '========================================================================
        ' Rows 61-64: Manual tech input only - no automation
        '========================================================================

        Case "$G$61", "$H$61", "$G$62", "$H$62", "$G$63", "$H$63", "$G$64", "$H$64"
            ' Manual input cells - just move to next cell
            On Error Resume Next
            ActiveCell.Offset(1, 0).Select
            If Err.Number <> 0 Then
                MsgBox "Error selecting next cell: " & Err.Description
                Err.Clear
            End If
            On Error GoTo 0

        '-------------------End Here on last Gray Cell---------------------
        Case "$G$65", "$H$65"
            CommToggle "Standby"
            Comm False, True, False
            ActiveSheet.Range("I9").Select
            TestSect = 0
            PrevSameTest = 0
            MsgBox "Verification Complete! Remove All Connections!"

        Case Else
            '-------------------Clicking anywhere else---------------------
            CommToggle "Standby"
            Comm False, True, False
            TestSect = 0
            PrevSameTest = 0
            SameTest = 0

    End Select

Else
    ' If status is not Standby or Operating, reset
    CommToggle "Standby"
    Comm False, True, False
    TestSect = 0
    PrevSameTest = 0
    SameTest = 0
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
