Attribute VB_Name = "Centered"
Sub CenterUserFormOnActiveSheet(UserForm As Object)
    Dim ws As Worksheet
    Dim ExcelLeft As Double
    Dim ExcelTop As Double
    Dim LeftPos As Double
    Dim TopPos As Double
    Dim TempWidth As Single
    Dim TempHeight As Single
    
    ' Get the active worksheet
    Set ws = ActiveSheet
    
    ' Get the position of the Excel application window
    ExcelLeft = Application.Left
    ExcelTop = Application.Top
    
    ' Calculate the left and top positions to center the userform
    LeftPos = ExcelLeft + (Application.Width - UserForm.Width) / 2
    TopPos = ExcelTop + (Application.Height - UserForm.Height) / 2
    
    ' Set the userform's position
    With UserForm
        .StartUpPosition = 0 ' Manually set the position
        .Left = LeftPos
        .Top = TopPos
    End With
End Sub

Sub DockPanelOnRight(UserForm As Object)
    Dim ExcelLeft As Double
    Dim ExcelTop As Double
    Dim RightPos As Double
    Dim TopPos As Double
    
    ' Get the position of the Excel application window
    ExcelLeft = Application.Left
    ExcelTop = Application.Top
    
    ' Top: align a bit below the title bar
    TopPos = ExcelTop + 50
    
    ' Left: dock to right edge of Excel window
    RightPos = ExcelLeft + Application.Width - UserForm.Width - 20
    
    ' Set the userform's position
    With UserForm
        .StartUpPosition = 0 ' Manual positioning
        .Left = RightPos
        .Top = TopPos
    End With
End Sub
