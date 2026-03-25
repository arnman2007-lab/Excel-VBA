Attribute VB_Name = "GValues"
Sub GetValues()
    Dim i As Long
    Dim startRow As Long
    Dim endRow As Long
    Dim currentRow As Long
    Dim Index As Integer
    Dim MaxIndex As Integer
    
    ArraySetup ' Ensure arrays are initialized
    
    ' Get current row
    currentRow = ActiveCell.Row

    ' Loop through range pairs
    For i = LBound(ranges) To UBound(ranges)
        ' Ensure range format is correct
        If InStr(1, ranges(i), ":") = 0 Then
            MsgBox "Error: Invalid range format in ranges(" & i & ") - " & ranges(i)
            Exit Sub
        End If

        ' Extract start and end rows
        startRow = CLng(Split(ranges(i), ":")(0))
        endRow = CLng(Split(ranges(i), ":")(1))

        ' Calculate the max index dynamically
        MaxIndex = endRow - startRow
        
        ' Check if current row is within range
        If currentRow >= startRow And currentRow <= endRow Then
            Index = currentRow - startRow

            ' Ensure Index is within bounds
            If Index < 0 Or Index > MaxIndex Then
                MsgBox "Error: Index out of bounds! Index = " & Index
                Exit Sub
            End If

            ' Assign values with checks for empty arrays
            If IsArray(TestPoint) And UBound(TestPoint) >= 0 Then
                OffValueV = TestPoint(Index)
            Else
                OffValueV = 0 ' Default value if array is empty
            End If

            If IsArray(TestPointUnits) And UBound(TestPointUnits) >= 0 Then
                OffValueU = TestPointUnits(Index)
            Else
                OffValueU = "" ' Default value if array is empty
            End If

            ' Check if TestPointFrequency is initialized and not empty
            If IsArray(TestPointFrequency) And UBound(TestPointFrequency) >= 0 Then
                OffValueHz = TestPointFrequency(Index)
                OffValueHzU = TestPointFrequencyUnits(Index)
            Else
                OffValueHz = 0
                OffValueHzU = ""
            End If
            
            ' Check if other arrays are initialized and not empty
            If IsArray(TestPointOffset) And UBound(TestPointOffset) >= 0 Then
                offset = TestPointOffset(Index)
            Else
                offset = 0
            End If
            
            If IsArray(TestPointComp) And UBound(TestPointComp) >= 0 Then
                OffSetU = TestPointComp(Index)
            Else
                OffSetU = 0
            End If
            
            If IsArray(TestPointWave) And UBound(TestPointWave) >= 0 Then
                Wave = TestPointWave(Index)
            Else
                Wave = ""
            End If
            
            If IsArray(TestPointDuty) And UBound(TestPointDuty) >= 0 Then
                duty = TestPointDuty(Index)
            Else
                duty = 0
            End If

            ' Debug values
            'MsgBox "OffValueV: " & OffValueV & " OffValueU: " & OffValueU & _
                   " OffValueHz: " & OffValueHz & " OffValueHzU: " & OffValueHzU & _
                   " Offset: " & offset & " OffSetU: " & OffSetU & " Wave: " & Wave & " Duty: " & duty
            
            Exit Sub ' Exit after finding the first match
        End If
    Next i
End Sub



'TestPoint(Index)
'TestPointUnits(Index)
'TestPointFrequency(Index)
'TestPointFrequencyUnits(Index)
'TestPointWave(Index)
'TestPointOffset(Index)
'TestPointComp(Index)
'TestPointDuty(Index)

Function GetArrayValue(arrayType As String, sectionNum As Integer, ByVal Index As Long) As Variant
    Dim arr As Variant
    Dim arrayName As String
    
    ' Construct array name dynamically (e.g., "TestPoint")
    arrayName = "WSSetup.Section" & sectionNum & arrayType
    
    ' Get the array reference directly from WSSetup module
    arr = CallByName(WSSetup, "Section" & sectionNum & arrayType, VbGet)
    
    ' Check if the array is valid and within bounds
    If Not IsError(arr) Then
        If Index >= LBound(arr) And Index <= UBound(arr) Then
            GetSectionArray = arr(Index)
        Else
            ' Return an empty value (or specific default) instead of Null
            GetSectionArray = ""  ' Can return "" or other default values instead of Null
        End If
    Else
        ' Return an empty value or default when the array does not exist
        GetSectionArray = ""  ' Can return "" or other default values instead of Null
    End If
End Function



