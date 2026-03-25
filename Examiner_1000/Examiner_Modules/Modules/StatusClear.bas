Attribute VB_Name = "StatusClear"

Sub ClearStatus()
    Dim ioMgr As VisaComLib.ResourceManager
    Dim instrument As VisaComLib.FormattedIO488
    
    Dim addstdvalue As Range
    
    Set addstdvalue = Worksheets("Information").Range("Calibrator")
    
If Range("Calibrator") = "" Then
    
    Else
    
    Set ioMgr = New VisaComLib.ResourceManager
    Set instrument = New VisaComLib.FormattedIO488
    Set instrument.IO = ioMgr.Open("GPIB0::" & addstdvalue.Value)
    
        
        instrument.WriteString "*cls"
        
        End If
End Sub
