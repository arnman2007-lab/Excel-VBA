Attribute VB_Name = "GetInformation"
Sub LoadDeviceInfoFromCSV()
    Dim fso As Object, ts As Object
    Dim filePath As String, line As String
    Dim Data() As String

    filePath = ThisWorkbook.Path & "\DeviceInfo.csv"

    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(filePath) Then
        MsgBox "DeviceInfo.csv not found.", vbExclamation
        Exit Sub
    End If

    Set ts = fso.OpenTextFile(filePath, 1)
    ts.SkipLine ' skip header line

    Do While Not ts.AtEndOfStream
        line = ts.ReadLine
        Data = Split(line, ",")
        
        Select Case Data(0)
            Case "Calibrator"
                WorkOrderSheet.Range("M8").Value = Data(1)
                WorkOrderSheet.Range("M9").Value = Data(2)
                WorkOrderSheet.Range("M10").Value = Data(3)
                WorkOrderSheet.Range("M11").Value = Data(4)
                WorkOrderSheet.Range("M12").Value = Data(5)
            Case "DMM"
                WorkOrderSheet.Range("P8").Value = Data(1)
                WorkOrderSheet.Range("P9").Value = Data(2)
                WorkOrderSheet.Range("P10").Value = Data(3)
                WorkOrderSheet.Range("P11").Value = Data(4)
            Case "Counter"
                WorkOrderSheet.Range("M15").Value = Data(1)
                WorkOrderSheet.Range("M16").Value = Data(2)
                WorkOrderSheet.Range("M17").Value = Data(3)
                WorkOrderSheet.Range("M18").Value = Data(4)
        End Select
    Loop

    ts.Close
End Sub

