Attribute VB_Name = "Module1"
Sub ListShapesAndButtons()
    Dim shp As Shape
    Dim btn As Button
    
    Debug.Print "=== Shapes on wsInfo ==="
    For Each shp In Sheets("Information").Shapes
        Debug.Print shp.Name, shp.Type
    Next shp
    
    Debug.Print "=== Buttons on wsInfo ==="
    For Each btn In Sheets("Information").Buttons
        Debug.Print btn.Name
    Next btn
End Sub

Sub ListAllShapesAndButtons()
    Dim ws As Worksheet
    Dim shp As Shape
    Dim btn As Button
    Dim ole As OLEObject
    
    For Each ws In ThisWorkbook.Worksheets
        Debug.Print "=== Sheet: " & ws.Name & " ==="
        
        ' List Shapes (pictures, form buttons, etc.)
        For Each shp In ws.Shapes
            Debug.Print " Shape: " & shp.Name & " | Type: " & shp.Type
        Next shp
        
        ' List Form Control buttons
        For Each btn In ws.Buttons
            Debug.Print " Form Button: " & btn.Name
        Next btn
        
        ' List ActiveX controls
        For Each ole In ws.OLEObjects
            Debug.Print " ActiveX Control: " & ole.Name & " | Type: " & ole.progID
        Next ole
    Next ws
End Sub

