Attribute VB_Name = "HVShow"
Sub ShowHVImageAtCell(targetCell As Range)
    Dim img As Shape
    Dim infoSheet As Worksheet
    Dim targetSheet As Worksheet
    Dim DataSheetName As String

    ' Dynamically assign DataSheet name (Tab1 or whatever you assign in SetupWS)
    'DataSheetName = "Tab1" ' Replace with your dynamic value, if applicable

    ' Set the sheets
    Set infoSheet = ThisWorkbook.Sheets("Information")
    Set targetSheet = ThisWorkbook.Sheets(Tab1) ' Use the dynamic sheet name

    On Error Resume Next
    Set img = targetSheet.Shapes("HVImage")
    On Error GoTo 0

    ' If the image doesn't exist on the target sheet, copy it from Information sheet
    If img Is Nothing Then
        infoSheet.Shapes("HVImage").Copy
        targetSheet.Paste
        Set img = targetSheet.Shapes(targetSheet.Shapes.Count) ' Get the last added shape (the image)
        img.Name = "HVImage"
    End If

    ' Move and show the image in the target cell
    With img
        .Top = targetCell.Top
        .Left = targetCell.Left
        .Visible = True
    End With
End Sub

Sub HVImageShow()
    Dim img As Shape
    Dim wsData As Worksheet
    Dim wsInfo As Worksheet
    Dim targetCell As Range

    Set wsInfo = ThisWorkbook.Sheets("Information")
    Set wsData = ThisWorkbook.Sheets(Tab1) ' Use variable Tab1

    On Error Resume Next
    Set img = wsData.Shapes("HVImage")
    On Error GoTo 0

    If img Is Nothing Then
        On Error Resume Next
        wsInfo.Shapes("HVImage").Copy
        wsData.Paste
        Set img = wsData.Shapes(wsData.Shapes.Count)
        img.Name = "HVImage"
        On Error GoTo 0
    End If

    img.Visible = msoFalse ' Always hide first

    ' Show only if conditions are met
    If OffValueU = "V" Then
        If OffValueV >= 100 Or OffValueV <= -100 Then
            Set targetCell = ActiveCell.offset(-7, 2)
            With img
                .Top = targetCell.Top
                .Left = targetCell.Left
                .Visible = msoTrue
            End With
        End If
    End If
End Sub



