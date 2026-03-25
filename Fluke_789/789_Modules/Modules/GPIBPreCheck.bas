Attribute VB_Name = "GPIBPreCheck"
Sub VerifyConnectedGPIBDevices()
    On Error Resume Next

    Dim resources() As String
    Dim RM_P As VisaComLib.ResourceManager
    Dim resourceName As Variant
    Dim detectedGPIB As Collection
    Dim InfoSheet As Worksheet
    Dim missingDevices As Boolean
    Dim foundAny As Boolean
    
    ' Cells storing expected device addresses
    Dim expectedCells As Variant
    expectedCells = Array("M11", "M18", "P11") ' Calibrator, Counter, DMM

    Set InfoSheet = ThisWorkbook.Sheets("Sheet1") ' Change to your actual sheet
    Set RM_P = New VisaComLib.ResourceManager
    Set detectedGPIB = New Collection

    resources = RM_P.FindRsrc("?*")

    ' Collect all GPIB device addresses
    For Each resourceName In resources
        If InStr(1, resourceName, "GPIB") > 0 And InStr(1, resourceName, "INTFC") = 0 Then
            detectedGPIB.Add resourceName
            foundAny = True
        End If
    Next resourceName

    Set RM_P = Nothing

    If Not foundAny Then
        ' No GPIB devices found at all — clear expected cells
        wsInfo.Range("M11").Value = ""
        wsInfo.Range("M18").Value = ""
        wsInfo.Range("P11").Value = ""
        MsgBox "No GPIB devices found. Device cells have been cleared.", vbCritical
        Exit Sub
    End If

    ' Check each expected device address is present in detected list
    Dim i As Long, expectedValue As String
    For i = LBound(expectedCells) To UBound(expectedCells)
        expectedValue = Trim(InfoSheet.Range(expectedCells(i)).Value)
        If expectedValue <> "" Then
            If Not CollectionContains(detectedGPIB, expectedValue) Then
                missingDevices = True
                Exit For
            End If
        End If
    Next i

    If missingDevices Then
        MsgBox "Some configured GPIB devices are missing or not responding." & vbCrLf & _
               "Please check connections. Re-run WorkStationSetup only if the setup has changed.", vbExclamation
    End If
End Sub

Private Function CollectionContains(col As Collection, key As String) As Boolean
    Dim item As Variant
    For Each item In col
        If item = key Then
            CollectionContains = True
            Exit Function
        End If
    Next item
    CollectionContains = False
End Function


