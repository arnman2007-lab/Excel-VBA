Attribute VB_Name = "OperationalChecks"
Sub OpChecks()

If Check = "Display" Then
'Display
ShowImageInCell "HVImage", "AA1"
            OperationChecks.show
            If TerminateClicked = True Then
            TerminateClicked = False
            Exit Sub
            End If
            If PassClicked = True Then
            ActiveCell = "Pass"
            ElseIf FailClicked = True Then
            ActiveCell = "Fail"
            End If
            Selection.offset(1, 0).Select
           
ElseIf Check = "Keypad" Then
            ShowImageInCell "HVImage", "AA1"
            OperationChecks.show
            If TerminateClicked = True Then
            TerminateClicked = False
            Exit Sub
            End If
            If PassClicked = True Then
            ActiveCell = "Pass"
            ElseIf FailClicked = True Then
            ActiveCell = "Fail"
            End If
           ' PrevAddress = activeCell.Address
            Selection.offset(1, 0).Select
            
           'Exit Sub
            
ElseIf Check = "Backlight" Then
            'MsgBox "in backlight"
            ShowImageInCell "HVImage", "AA1"
            OperationChecks.show
            If TerminateClicked = True Then
            TerminateClicked = False
            Exit Sub
            End If
            If PassClicked = True Then
            ActiveCell = "Pass"
            ElseIf FailClicked = True Then
            ActiveCell = "Fail"
            End If
            Selection.offset(1, 0).Select
                        
ElseIf Check = "CurrentSense" Then
            ShowImageInCell "HVImage", "AA1"
            OperationChecks.show
            If TerminateClicked = True Then
            TerminateClicked = False
            Exit Sub
            End If
            If PassClicked = True Then
            ActiveCell = "Pass"
            ElseIf FailClicked = True Then
            ActiveCell = "Fail"
            End If
            Selection.offset(1, 0).Select
            
ElseIf Check = "PowerLED" Then
            ShowImageInCell "HVImage", "AA1"
            OperationChecks.show
            If TerminateClicked = True Then
            TerminateClicked = False
            Exit Sub
            End If
            If PassClicked = True Then
            ActiveCell = "Pass"
            ElseIf FailClicked = True Then
            ActiveCell = "Fail"
            End If
            Selection.offset(1, 0).Select
End If
PrevAddress = ActiveCell.Address
End Sub
