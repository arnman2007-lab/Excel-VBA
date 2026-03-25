Attribute VB_Name = "HVShow"


Sub HVImageShow(ParamC As Double, ParamCUnit As String)
    Dim wsData As Worksheet
    Dim wsInfo As Worksheet
    Dim targetCell As Range
    Dim img As Shape
    Dim imagePath As String
    
    SetupWS
    Set wsData = ThisWorkbook.Sheets(Tab1)
    Set wsInfo = ThisWorkbook.Sheets("Information")
    
    ' Full path to the image
    imagePath = ThisWorkbook.Path & "\Images\HVImage.jpg"
    
    ' Delete any existing HVImage on the datasheet first
    On Error Resume Next
    wsData.Shapes("HVImage").Delete
    On Error GoTo 0
    
    ' If condition is met, insert image at target location
    If ParamC >= 100 Or ParamC <= -100 And ParamCUnit = "V" Then
   ' MsgBox "inside hvimge" & ParamC
        Set targetCell = ActiveCell.OffSet(-7, 2)
        
        Set img = wsData.Shapes.AddPicture( _
                    Filename:=imagePath, _
                    LinkToFile:=msoFalse, _
                    SaveWithDocument:=msoCTrue, _
                    Left:=targetCell.Left, _
                    Top:=targetCell.Top, _
                    Width:=-1, _
                    Height:=-1)
        img.Name = "HVImage"
        
    Else
   
    End If
End Sub


