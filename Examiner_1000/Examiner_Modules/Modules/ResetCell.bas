Attribute VB_Name = "ResetCell"
Sub ResetCells(tabName As String)
    SetupWS
    If Not IsArray(ranges) Or IsEmpty(ranges) Then
        ArraySetup
    End If

    Dim i As Integer
    Dim j As Long
    Dim rangePair As Variant
    Dim startRow As Long
    Dim endRow As Long
    Dim rangeSplit() As String
    Dim afCell As Range, alCell As Range

    If tabName = Tab1 Then
        i = 0
        
        ' Information Tab
        WorkOrderSheet.Range("H13").Value = ""
        i = i + 1
        WorkOrderSheet.Range("X3").Value = Make
        i = i + 1
        WorkOrderSheet.Range("Y3").Value = Model
        i = i + 1
        WorkOrderSheet.Range("W4").Value = UnitDesc
        i = i + 1
        WorkOrderSheet.Range("H14").Value = ""
        i = i + 1
        WorkOrderSheet.Range("H15").Value = ""
        i = i + 1
        WorkOrderSheet.Range("H16").Value = "N/A"
        i = i + 1

        ' Main loop through ranges
        For Each rangePair In ranges
            rangeSplit = Split(rangePair, ":")
            startRow = CLng(rangeSplit(0))
            endRow = CLng(rangeSplit(1))
            
            For j = startRow To endRow
                Set afCell = dataSheet.Range(ColLetterAF & j)
                If Not afCell.MergeCells Or afCell.Address = afCell.MergeArea.Cells(1, 1).Address Then
                    afCell.Value = ""
                    i = i + 1
                End If
                
                Set alCell = dataSheet.Range(ColLetterAL & j)
                If Not alCell.MergeCells Or alCell.Address = alCell.MergeArea.Cells(1, 1).Address Then
                    alCell.Value = "N/A"
                    i = i + 1
                End If
            Next j
        Next rangePair

    ElseIf tabName = Tab2 Then
        ' Handle Tab2 logic here (if any)
        
    ElseIf tabName = Tab3 Then
        ' Handle Tab3 logic here (if any)
        
    ElseIf tabName = Tab4 Then
        ' Handle Tab4 logic here (if any)

    End If
End Sub


Sub Inop(tabName As String)
SetupWS
Dim i As Integer
Dim j As Long
Dim rangePair As Variant
    Dim startRow As Long
    Dim endRow As Long
    Dim rangeSplit() As String



If tabName = Tab1 Then
i = 0
    'Information Tab
    'For j = 13 To 15
    WorkOrderSheet.Range("H13").Value = ""
    i = i + 1
    WorkOrderSheet.Range("X3").Value = Make
    i = i + 1
    WorkOrderSheet.Range("Y3").Value = Model
    i = i + 1
    WorkOrderSheet.Range("W4").Value = UnitDesc
    i = i + 1
    WorkOrderSheet.Range("H14").Value = ""
    i = i + 1
    WorkOrderSheet.Range("H15").Value = ""
    i = i + 1
    WorkOrderSheet.Range("H16").Value = "N/A"
    i = i + 1
    
    
    
For Each rangePair In ranges
        
        rangeSplit = Split(rangePair, ":")
        startRow = CLng(rangeSplit(0))
        endRow = CLng(rangeSplit(1))
    For j = startRow To endRow
    dataSheet.Range("F" & j).Value = "INOP"
    i = i + 1
    dataSheet.Range("G" & j).Value = "N/A"
    i = i + 1
    Next j
    Next rangePair

ElseIf tabName = Tab2 Then
ElseIf tabName = Tab3 Then
ElseIf tabName = Tab4 Then





End If
End Sub

