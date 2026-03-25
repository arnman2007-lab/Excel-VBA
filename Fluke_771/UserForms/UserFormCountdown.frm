VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormCountdown 
   Caption         =   "UserForm1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "UserFormCountdown.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserFormCountdown"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private countdownSeconds As Integer

Private Sub UserForm_Activate()
 CenterUserFormOnActiveSheet Me
    countdownSeconds = 15
    UpdateCountdown
    ' Start timer loop
    DoEvents
    StartTimer
End Sub

Private Sub StartTimer()
    Dim startTime As Single
    startTime = Timer

    Do While countdownSeconds > 0
        DoEvents
        If Timer - startTime >= 1 Then
            countdownSeconds = countdownSeconds - 1
            startTime = Timer
            UpdateCountdown
        End If
    Loop
    
    Unload Me
End Sub

Private Sub UpdateCountdown()
    Me.lblCountdown.Caption = "Taking readings for " & countdownSeconds & " seconds..." & vbCrLf & _
                              "DMM is taking 30 readings. Please wait."
End Sub


