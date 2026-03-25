Attribute VB_Name = "DMMStatusClear"

Sub DMMClearStatus(StatusClear As String)
    
    DMMGPIB = wsInfo.Range("$P$11").Value
    DMMModel = wsInfo.Range("$P$9").Value
    
'Exit sub if no gpib address
If DMMGPIB = "" Then Exit Sub

    Set ioMgr = New VisaComLib.ResourceManager
    Set DMMDevice = New VisaComLib.FormattedIO488
    Set DMMDevice.IO = ioMgr.Open(DMMGPIB)

Select Case StatusClear

    Case "Close"
    
        Select Case DMMModel

            Case "3458A"
                PanelForm.STDAction.Caption = DMMModel & " Send Command: Close"
                DoEvents
                Set gpib3458 = ioMgr.Open(DMMDevice.IO.resourceName)
                DMMDevice.WriteString "RESET"
                gpib3458.ControlREN GPIB_REN_GTL
                gpib3458.Close
                
        
            Case "8508A"
                PanelForm.STDAction.Caption = DMMModel & " Send Command: Close"
                DoEvents
                DMMDevice.WriteString "Reset"
                
        
            Case "34401A"
                PanelForm.STDAction.Caption = DMMModel & " Send Command: Close"
                DoEvents
                DMMDevice.WriteString "Reset"
                

        
        End Select
    
    Case "Clear"
    
        Select Case DMMModel

            Case "3458A"
                PanelForm.STDAction.Caption = DMMModel & " Send Command: Clear"
                DoEvents
                DMMDevice.WriteString "Reset"
                
        
            Case "8508A"
                PanelForm.STDAction.Caption = DMMModel & " Send Command: Clear"
                DoEvents
                DMMDevice.WriteString "Reset"
                
        
            Case "34401A"
                PanelForm.STDAction.Caption = DMMModel & " Send Command: Clear"
                DoEvents
                DMMDevice.WriteString "Reset"
                

        
        End Select
        
    Case "Reset"
    
        Select Case DMMModel

            Case "3458A"
                PanelForm.STDAction.Caption = DMMModel & " Send Command: Reset"
                DoEvents
                DMMDevice.WriteString "Reset"
                
        
            Case "8508A"
                PanelForm.STDAction.Caption = DMMModel & " Send Command: Reset"
                DoEvents
                DMMDevice.WriteString "Reset"
                
        
            Case "34401A"
                PanelForm.STDAction.Caption = DMMModel & " Send Command: Reset"
                DoEvents
                DMMDevice.WriteString "Reset"
                

        
        End Select
        
        
    Case "Standby"
    
        Select Case DMMModel

            Case "3458A"
                DMMDevice.WriteString "Reset"
                
        
            Case "8508A"
                DMMDevice.WriteString "Reset"
                
        
            Case "34401A"
                DMMDevice.WriteString "Reset"
                

        
        End Select
    
    
    
End Select
    
    'Else
    

    
        
       ' instrument.WriteString "*cls"
        
       ' End If
End Sub

