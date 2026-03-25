VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} OperationChecks 
   Caption         =   "Is it on?"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "OperationChecks.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "OperationChecks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





Private Sub NoCont_Click()

ActiveCell = "Fail"
Unload Me


End Sub

Private Sub Terminate_Click()
    TerminateClicked = True
    ActiveCell.offset(0, -2).Select
    PrevAddress = ActiveCell.Address
    shouldContinue = False

    
Unload Me
    
End Sub



Private Sub YesCont_Click()

ActiveCell = "Pass"
Unload Me


End Sub

Private Sub UserForm_Initialize()
 CenterUserFormOnActiveSheet Me
Call CloseButtonSettings(Me, False)
SetupWS
 Dim imagePath As String
    Dim CalibModel As String
    
    
    
   
    
        If Check = "Display" Then
            Me.Caption = dataSheet.Range("B14").Value
            Me.Label1.Caption = "Connect to " & CalibModel & " to Read AC Voltage"
            
        ElseIf Check = "Keypad" Then
            Me.Caption = dataSheet.Range("B16").Value
            Me.Label1.Caption = "All buttons beep when pressed."
            
        ElseIf Check = "Backlight" Then
            Me.Caption = dataSheet.Range("B15").Value
            Me.Label1.Caption = "Two brightness levels."
            
        ElseIf Check = "Display" Then
            Me.Caption = dataSheet.Range("B17").Value
            Me.Label1.Caption = "Connect to " & CalibModel & " to Read DC Hz"
            
        ElseIf Check = "CurrentSense" Then
            Me.Caption = dataSheet.Range("B17").Value
            Me.Label1.Caption = "Alarm sounds when mA/µA and A jacks have a lead plug inserted."
            
        ElseIf Check = "PowerLED" Then
            Me.Caption = dataSheet.Range("B14").Value
            Me.Label1.Caption = "Lights for 4 seconds on power-up then goes off"
            
        ElseIf Check = "None" Then
            Me.Caption = dataSheet.Range("B56").Value
            Me.Label1.Caption = "Connect to " & CalibModel & " to Read Capacitance"
             
        ElseIf Check = "None" Then
            Me.Caption = dataSheet.Range("B59").Value
            Me.Label1.Caption = "Connect to " & CalibModel & " to Test Continuity"
            
        ElseIf Check = "None" Then
            Me.Caption = dataSheet.Range("B61").Value
            Me.Label1.Caption = "Connect to " & CalibModel & " to Test Diode"
            
        ElseIf Check = "None" Then
            Me.Caption = dataSheet.Range("B63").Value
            Me.Label1.Caption = "Connect to " & CalibModel & " to Read AC MilliAmps"
            
        End If
            


End Sub

