Attribute VB_Name = "TestSectionTests"
Sub TestOp()
ArraySetup
'MsgBox TestSect
'MsgBox "SameTest: " & SameTest & "PrevSameTest: " & PrevSameTest
    'If SameTest = 1 Then
    'PrevSameTest = SameTest
    'End If
    'If you have a TestSect that needs special parameters then you can make a If statement for that TestSect in the IF statement Below
    
    If PrevSameTest = SameTest Then
    
    Out
Else
    MainHookup.show
    If TerminateClicked Then
    TerminateClicked = False
    Exit Sub
End If

    
    

    Out
    
    'PrevSameTest = SameTest
End If
PrevSameTest = SameTest
End Sub


