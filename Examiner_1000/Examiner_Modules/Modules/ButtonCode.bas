Attribute VB_Name = "ButtonCode"
Public Sub DSPrint_Click()
    SetupWS
        
If Range("J8").Value = "Status: Incomplete" Or WorkOrderSheet.Range("H14").Value = "" Or WorkOrderSheet.Range("H15").Value = "" Or WorkOrderSheet.Range("H16").Value = "" Then
MsgBox "Please Fix Empty Cells - Check spaces between input cells too"
        
Exit Sub
End If
        
    PrintSelection.show

End Sub

Public Sub ResetDatasheet_Click()
    Dim tabName     As String
    tabName = ActiveSheet.Name
    
    ResetCells tabName
    'RunExternalCode "ResetCell", "ResetCells", tabName
End Sub

Public Sub GetData_Click()
    Dim tabName     As String
    tabName = ActiveSheet.Name
    'RunExternalCode "DataSave", "RetrieveData", tabName
    RetrieveData tabName
End Sub

Public Sub SetINOP_Click()
    Dim tabName As String
    tabName = ActiveSheet.Name
    'RunExternalCode "ResetCell", "Inop", tabName
    Inop tabName
End Sub

Public Sub StoreInputData_Click()
    Dim tabName     As String
    tabName = ActiveSheet.Name
    'RunExternalCode "DataSave", "StoreData", tabName
    StoreData tabName
    
End Sub

Public Sub PreloadStuff_Click()
    Dim tabName     As String
    tabName = ActiveSheet.Name
    'RunExternalCode "PreloadCells", "Preload", tabName
    Preload tabName
End Sub

Public Sub ToggleButton1_Click()

With Sheet2
        If .Range("AA1").Value = "OFF" Then
            ' Turn OFF
            
            CommToggle "Standby"
        Else
            ' Turn Standby
            
            CommToggle "OFF"
            
        End If
    End With
End Sub
Public Sub CommToggle(ToggleState As String)
    Dim btn As Shape
    Set btn = Sheet2.Shapes("CommToggle")
'MsgBox "CommToggle " & ToggleState
    
        If ToggleState = "OFF" Then
            ' Turn OFF
            
            Sheet2.Range("AA1").Value = "OFF"
            With btn
                .TextFrame.Characters.Text = "OFF"
                .TextFrame.Characters.Font.Color = RGB(0, 0, 0) ' black text
                .Fill.ForeColor.RGB = RGB(255, 255, 0) ' yellow background
            End With
            Comm True, False, False
        ElseIf ToggleState = "Standby" Then
            ' Turn Standby
            
            Sheet2.Range("AA1").Value = "Standby"
            With btn
                .TextFrame.Characters.Text = "Standby"
                .TextFrame.Characters.Font.Color = RGB(0, 0, 255) ' black text
                .Fill.ForeColor.RGB = RGB(0, 0, 0) ' yellow background
                .Fill.BackColor.RGB = RGB(0, 0, 255)
            End With
            Comm True, False, False
        ElseIf ToggleState = "Operating" Then
            ' Turn ON
            Sheet2.Range("AA1").Value = "Operating"
            
            With btn
                .TextFrame.Characters.Text = "Operating"
                .TextFrame.Characters.Font.Color = RGB(255, 0, 0) ' white text
                .TextFrame.Characters.Font.Size = 20
                .TextFrame.Characters.Font.Bold = True
                .TextFrame.Characters.Font.Color = RGB(255, 0, 0)
                    
                    
                
            End With
            Comm False, False, True
            
        End If
   
End Sub

