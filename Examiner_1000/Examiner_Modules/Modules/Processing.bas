Attribute VB_Name = "Processing"
Sub ProcessTestPoints()
    Dim sectionNumber As Integer
    Dim testPointValue As Variant
    Dim TestPointUnits As Variant
    Dim frequencyValue As Variant
    Dim frequencyUnits As Variant
    Dim waveform As Variant
    Dim offset As Variant
    Dim comp As Variant
    Dim duty As Variant
    
    ' Get current row
    Dim currentRow As Long
    currentRow = ActiveCell.Row
    
    ' Check which range the current row falls into
    For sectionNumber = 1 To 30
        ' Dynamically build variable names
        testPointValue = Eval("Section" & sectionNumber & "TestPoint(0)")
        TestPointUnits = Eval("Section" & sectionNumber & "TestPointUnits(0)")
        frequencyValue = Eval("Section" & sectionNumber & "TestPointFrequency(0)")
        frequencyUnits = Eval("Section" & sectionNumber & "TestPointFrequencyUnits(0)")
        waveform = Eval("Section" & sectionNumber & "TestPointWave(0)")
        offset = Eval("Section" & sectionNumber & "TestPointOffset(0)")
        comp = Eval("Section" & sectionNumber & "TestPointComp(0)")
        duty = Eval("Section" & sectionNumber & "TestPointDuty(0)")
        
        ' Check if this section matches the current row
        If isInRange(currentRow, sectionNumber) Then
            ' Process the test point if it's not Null
            If Not IsNull(testPointValue) Then
                ProcessTestPoint _
                    testPointValue, _
                    TestPointUnits, _
                    frequencyValue, _
                    frequencyUnits, _
                    waveform, _
                    offset, _
                    comp, _
                    duty
                
                ' Move to next cell
                ActiveCell.offset(1, 0).Select
                Exit Sub
            End If
        End If
    Next sectionNumber
End Sub



Sub ProcessTestPoint( _
    testPointValue As Variant, _
    TestPointUnits As Variant, _
    frequencyValue As Variant, _
    frequencyUnits As Variant, _
    waveform As Variant, _
    offsetValue As Variant, _
    comp As Variant, _
    dutyValue As Variant)
    
    ' Assign values to the specified variables
    OffValueV = testPointValue
    OffValueU = TestPointUnits
    OffValueHz = frequencyValue
    OffValueHzU = frequencyUnits
    Wave = waveform
    offset = offsetValue
    comp = comp
    duty = dutyValue
    
    ' Optional: Debug print to verify assignments
    Debug.Print "Assigned Values:"
    Debug.Print "OffValueV: " & OffValueV & " " & OffValueU
    Debug.Print "OffValueHz: " & OffValueHz & " " & OffValueHzU
    Debug.Print "Wave: " & Wave
    Debug.Print "Offset: " & offset
    Debug.Print "Comp: " & comp
    Debug.Print "Duty: " & duty
End Sub

Function GetArrayIndexForRow(currentRow As Long) As Long
    Dim i As Long
    Dim rowRange As Variant
    Dim startRow As Long
    Dim endRow As Long
    
    ' Loop through the global ranges array
    For i = LBound(ranges) To UBound(ranges)
        ' Split the range string into start and end rows
        rowRange = Split(ranges(i), ":")
        startRow = CLng(rowRange(0))
        endRow = CLng(rowRange(1))
        
        ' Check if the current row is within this range
        If currentRow >= startRow And currentRow <= endRow Then
            ' Calculate the index within the test point array
            ' Assumes the first row of the range corresponds to index 0
            GetArrayIndexForRow = currentRow - startRow
            Exit Function
        End If
    Next i
    
    ' Return -1 if no matching range is found
    GetArrayIndexForRow = -1
End Function
