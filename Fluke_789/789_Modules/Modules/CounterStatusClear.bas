Attribute VB_Name = "CounterStatusClear"

Sub CounterClearStatus(StatusClear As String)
    
    CounterGPIB = wsInfo.Range("$M$11").Value
    CounterModel = wsInfo.Range("$M$9").Value
'Exit sub if no gpib address
 
If CounterGPIB = "" Then Exit Sub

    'Set ioMgr = New VisaComLib.ResourceManager
   ' Set CounterDevice = New VisaComLib.FormattedIO488
    'Set CounterDevice.IO = ioMgr.Open(CounterGPIB)

Select Case StatusClear

    Case "Clear"
    
        Select Case CounterModel

            Case "5500A"
                CounterDevice.WriteString "*cls; *OPC?"
                ready = CounterDevice.ReadString()
        
            Case "5502A"
                CounterDevice.WriteString "*cls; *OPC?"
                ready = CounterDevice.ReadString()
        
            Case "5520A"
                CounterDevice.WriteString "*cls; *OPC?"
                ready = CounterDevice.ReadString()
                
        
            Case "5522A"
                CounterDevice.WriteString "*cls; *OPC?"
                ready = CounterDevice.ReadString()
        
            Case "M3001"
                CounterDevice.WriteString "*cls; *OPC?"
                ready = CounterDevice.ReadString()
        
        End Select
        
    Case "Reset"
    
        Select Case CounterModel

            Case "5500A"
                CounterDevice.WriteString "*RST; *OPC?"
                ready = CounterDevice.ReadString()
        
            Case "5502A"
                CounterDevice.WriteString "*RST; *OPC?"
                ready = CounterDevice.ReadString()
        
            Case "5520A"
                MsgBox "We Are Resetting"
                'CounterDevice.WriteString "*RST; *OPC?"
                'ready = CounterDevice.ReadString()
        
            Case "5522A"
                CounterDevice.WriteString "*RST; *OPC?"
                ready = CounterDevice.ReadString()
        
            Case "M3001"
                CounterDevice.WriteString "*RST; *OPC?"
                ready = CounterDevice.ReadString()
        
        End Select
        
        
    Case "Standby"
    
        Select Case CounterModel

            Case "5500A"
                'Counter "Source", "DCV", "", ""
                CounterDevice.WriteString "STBY; *OPC?"
                ready = CounterDevice.ReadString()
        
            Case "5502A"
                'Counter "Source", "DCV"
                CounterDevice.WriteString "STBY; *OPC?"
                ready = CounterDevice.ReadString()
        
            Case "5520A"
                'Counter "Source", "DCV"
                CounterDevice.WriteString "STBY; *OPC?"
                ready = CounterDevice.ReadString()
        
            Case "5522A"
                'Counter "Source", "DCV"
                CounterDevice.WriteString "STBY; *OPC?"
                ready = CounterDevice.ReadString()
        
            Case "M3001"
                'Counter "Source", "DCV"
                CounterDevice.WriteString "STBY; *OPC?"
                ready = CounterDevice.ReadString()
        
        End Select
    
    
    
End Select
    
    'Else
    

    
        
       ' instrument.WriteString "*cls"
        
       ' End If
End Sub

