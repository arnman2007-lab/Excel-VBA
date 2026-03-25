Attribute VB_Name = "EmptyCellCheck"
Sub CheckforEmptyCommPortCell()


    Set usbcommvalue = Range("O8")

    If IsEmpty(usbcommvalue) = False Then
   
        Application.Wait (Now + TimeValue("0:00:3"))
        'Call Fluke789Reading
        Application.Wait (Now + TimeValue("0:00:1"))
        Selection.offset(1, 0).Select
    Else
        'USBComm cell is empty
    End If
End Sub
