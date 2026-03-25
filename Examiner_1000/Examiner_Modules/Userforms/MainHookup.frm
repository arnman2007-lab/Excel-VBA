VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MainHookup 
   Caption         =   "Set Unit to measure Frequency"
   ClientHeight    =   9420.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9120.001
   OleObjectBlob   =   "MainHookup.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "MainHookup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False







Private Sub Advance_Click()
Unload Me
shouldContinue = True

End Sub

Private Sub Fail_Click()
ActiveCell.Value = "Fail"
Unload Me
Selection.offset(1, 0).Select
End Sub





Private Sub Pass_Click()
ActiveCell.Value = "Pass"
Unload Me
Selection.offset(1, 0).Select
End Sub

Private Sub Terminate_Click()
    TerminateClicked = True
    ActiveCell.offset(0, -2).Select
    PrevAddress = ActiveCell.Address
    shouldContinue = False

    
Unload Me
    
End Sub

Private Sub UserForm_Initialize()
 CenterUserFormOnActiveSheet Me
Call CloseButtonSettings(Me, False)
SetupWS
 Dim imagePath As String
 Dim STDStyle As Integer
    
   Me.Pass.Visible = False
   Me.Fail.Visible = False
   Me.Advance.Visible = True
   'Do not change this line
   Me.Label11.Caption = Model
   
    
    
        'Copy and paste the insides, (between If TestSect = 1 Then and ElseIf TestSect = 2 Then - just the code), of TestSect1 into the ElseIf TestSect = 2 Then Fill out
        
        If TestSect = 1 Then
            'This is the Function Description for example AC Voltage Tests @ 60 Hz - From Datasheet Just click the test description and get the cell address for example B14
            'Then change the address in the quotes below
            Me.Caption = dataSheet.Range("B19").Value
            'this is the path to the hookup image, if using Naming scheme(highly recommended) it will look up the correct image for the unit and Standard model numbers
            imagePath = ThisWorkbook.Path & "\Images\" & Model & "\" & CalibModel & "\A Main Hookup " & CalibModel & ".jpg"
            
            'This is the Title of the userForm at the very top
            Me.Label6.Caption = "Connect to " & CalibModel & " to Read Acceleration Response"
            'This is the main information for the test ie Turn Knob, push button, eat a snack etc.
            Me.Label33.Caption = "Press On/Select to Select A"
            
           
            'Me.Label13.Caption = DMMModel
            Me.Label35.Visible = False  '3458 Userform True to show False to not show
            
            'Unit to Calibrator/DMM Connections
            Me.Label37.Caption = "BNC"   'Wiring from Meter
            Me.Label31.Caption = "<---->"           '<---->  Hookup between Unit and standards
            Me.Label30.Caption = "Normal Hi"        'Wiring to Calibrator/DMM
            Me.Label29.Caption = "BNC"              'Wiring from Meter
            Me.Label28.Caption = "<---->"           '<---->  Hookup between Unit and standards
            Me.Label27.Caption = "Normal Lo"        'Wiring to Calibrator/DMM
            Me.Label26.Caption = ""                 'Wiring from Meter
            Me.Label25.Caption = ""                 '<---->  Hookup between Unit and standards
            Me.Label24.Caption = ""                 'Wiring to Calibrator/DMM
            Me.Label14.Caption = ""                 'Wiring from Meter
            Me.Label23.Caption = ""                 '<---->  Hookup between Unit and standards
            Me.Label22.Caption = ""                 'Wiring to Calibrator/DMM
            Me.Label34.Caption = ""  'Mid Userform Location between Label33 and Units jack inputs various information i.e. Turn knob, push button, wait till stable etc
            
        ElseIf TestSect = 2 Then
            'This is the Function Description for example AC Voltage Tests @ 60 Hz - From Datasheet Just click the test description and get the cell address for example B14
            'Then change the address in the quotes below
            Me.Caption = dataSheet.Range("B28").Value
            'this is the path to the hookup image, if using Naming scheme(highly recommended) it will look up the correct image for the unit and Standard model numbers
            imagePath = ThisWorkbook.Path & "\Images\" & Model & "\" & CalibModel & "\A Main Hookup " & CalibModel & ".jpg"
            
            'This is the Title of the userForm at the very top
            Me.Label6.Caption = "Connect to " & CalibModel & " to Read Acceleration Accuracy"
            'This is the main information for the test ie Turn Knob, push button, eat a snack etc.
            Me.Label33.Caption = "Press On/Select to Select A"
            Me.Label35.Visible = False  '3458 Userform True to show False to not show
            
            'Unit to Calibrator/DMM Connections
            Me.Label32.Caption = "BNC"              'Wiring from Meter
            Me.Label31.Caption = "<---->"           '<---->  Hookup between Unit and standards
            Me.Label30.Caption = "Normal Hi"        'Wiring to Calibrator/DMM
            Me.Label29.Caption = "BNC"              'Wiring from Meter
            Me.Label28.Caption = "<---->"           '<---->  Hookup between Unit and standards
            Me.Label27.Caption = "Normal Lo"        'Wiring to Calibrator/DMM
            Me.Label26.Caption = ""                 'Wiring from Meter
            Me.Label25.Caption = ""                 '<---->  Hookup between Unit and standards
            Me.Label24.Caption = ""                 'Wiring to Calibrator/DMM
            Me.Label14.Caption = ""                 'Wiring from Meter
            Me.Label23.Caption = ""                 '<---->  Hookup between Unit and standards
            Me.Label22.Caption = ""                 'Wiring to Calibrator/DMM
            Me.Label34.Caption = ""  'Mid Userform Location between Label33 and Units jack inputs various information i.e. Turn knob, push button, wait till stable etc
            
        ElseIf TestSect = 3 Then
            'This is the Function Description for example AC Voltage Tests @ 60 Hz - From Datasheet Just click the test description and get the cell address for example B14
            'Then change the address in the quotes below
            Me.Caption = dataSheet.Range("B39").Value
            'this is the path to the hookup image, if using Naming scheme(highly recommended) it will look up the correct image for the unit and Standard model numbers
            imagePath = ThisWorkbook.Path & "\Images\" & Model & "\" & CalibModel & "\V Main Hookup " & CalibModel & ".jpg"
            
            'This is the Title of the userForm at the very top
            Me.Label6.Caption = "Connect to " & CalibModel & " to Read Velocity Response"
            'This is the main information for the test ie Turn Knob, push button, eat a snack etc.
            Me.Label33.Caption = "Press On/Select to Select V"
            Me.Label35.Visible = False  '3458 Userform True to show False to not show
            
            'Unit to Calibrator/DMM Connections
            Me.Label32.Caption = "BNC"              'Wiring from Meter
            Me.Label31.Caption = "<---->"           '<---->  Hookup between Unit and standards
            Me.Label30.Caption = "Normal Hi"        'Wiring to Calibrator/DMM
            Me.Label29.Caption = "BNC"              'Wiring from Meter
            Me.Label28.Caption = "<---->"           '<---->  Hookup between Unit and standards
            Me.Label27.Caption = "Normal Lo"        'Wiring to Calibrator/DMM
            Me.Label26.Caption = ""                 'Wiring from Meter
            Me.Label25.Caption = ""                 '<---->  Hookup between Unit and standards
            Me.Label24.Caption = ""                 'Wiring to Calibrator/DMM
            Me.Label14.Caption = ""                 'Wiring from Meter
            Me.Label23.Caption = ""                 '<---->  Hookup between Unit and standards
            Me.Label22.Caption = ""                 'Wiring to Calibrator/DMM
            Me.Label34.Caption = ""  'Mid Userform Location between Label33 and Units jack inputs various information i.e. Turn knob, push button, wait till stable etc
            
        ElseIf TestSect = 4 Then
            'This is the Function Description for example AC Voltage Tests @ 60 Hz - From Datasheet Just click the test description and get the cell address for example B14
            'Then change the address in the quotes below
            Me.Caption = dataSheet.Range("B53").Value
            'this is the path to the hookup image, if using Naming scheme(highly recommended) it will look up the correct image for the unit and Standard model numbers
            imagePath = ThisWorkbook.Path & "\Images\" & Model & "\" & CalibModel & "\E Main Hookup " & CalibModel & ".jpg"
            
            'This is the Title of the userForm at the very top
            Me.Label6.Caption = "Connect to " & CalibModel & " to Read Acceleration Envelope"
            'This is the main information for the test ie Turn Knob, push button, eat a snack etc.
            Me.Label33.Caption = "Press On/Select to Select E"
            Me.Label35.Visible = False  '3458 Userform True to show False to not show
            
            Me.Pass.Visible = True
            Me.Fail.Visible = True
            Me.Advance.Visible = False
            
            'Unit to Calibrator/DMM Connections
            Me.Label32.Caption = "BNC"              'Wiring from Meter
            Me.Label31.Caption = "<---->"           '<---->  Hookup between Unit and standards
            Me.Label30.Caption = "Normal Hi"        'Wiring to Calibrator/DMM
            Me.Label29.Caption = "BNC"              'Wiring from Meter
            Me.Label28.Caption = "<---->"           '<---->  Hookup between Unit and standards
            Me.Label27.Caption = "Normal Lo"        'Wiring to Calibrator/DMM
            Me.Label26.Caption = ""                 'Wiring from Meter
            Me.Label25.Caption = ""                 '<---->  Hookup between Unit and standards
            Me.Label24.Caption = ""                 'Wiring to Calibrator/DMM
            Me.Label14.Caption = ""                 'Wiring from Meter
            Me.Label23.Caption = ""                 '<---->  Hookup between Unit and standards
            Me.Label22.Caption = ""                 'Wiring to Calibrator/DMM
            Me.Label34.Caption = "Is Acceleration Envelope <= 0.5 G"  'Mid Userform Location between Label33 and Units jack inputs various information i.e. Turn knob, push button, wait till stable etc
            
        ElseIf TestSect = 5 Then
            'This is the Function Description for example AC Voltage Tests @ 60 Hz - From Datasheet Just click the test description and get the cell address for example B14
            'Then change the address in the quotes below
            Me.Caption = dataSheet.Range("B53").Value
            'this is the path to the hookup image, if using Naming scheme(highly recommended) it will look up the correct image for the unit and Standard model numbers
            imagePath = ThisWorkbook.Path & "\Images\" & Model & "\" & CalibModel & "\E Main Hookup " & CalibModel & ".jpg"
            
            'This is the Title of the userForm at the very top
            Me.Label6.Caption = "Connect to " & CalibModel & " to Read Acceleration Envelope"
            'This is the main information for the test ie Turn Knob, push button, eat a snack etc.
            Me.Label33.Caption = "Press On/Select to Select E"
            Me.Label35.Visible = False  '3458 Userform True to show False to not show
            
            Me.Pass.Visible = True
            Me.Fail.Visible = True
            Me.Advance.Visible = False
            
            'Unit to Calibrator/DMM Connections
            Me.Label32.Caption = "BNC"              'Wiring from Meter
            Me.Label31.Caption = "<---->"           '<---->  Hookup between Unit and standards
            Me.Label30.Caption = "Normal Hi"        'Wiring to Calibrator/DMM
            Me.Label29.Caption = "BNC"              'Wiring from Meter
            Me.Label28.Caption = "<---->"           '<---->  Hookup between Unit and standards
            Me.Label27.Caption = "Normal Lo"        'Wiring to Calibrator/DMM
            Me.Label26.Caption = ""                 'Wiring from Meter
            Me.Label25.Caption = ""                 '<---->  Hookup between Unit and standards
            Me.Label24.Caption = ""                 'Wiring to Calibrator/DMM
            Me.Label14.Caption = ""                 'Wiring from Meter
            Me.Label23.Caption = ""                 '<---->  Hookup between Unit and standards
            Me.Label22.Caption = ""                 'Wiring to Calibrator/DMM
            Me.Label34.Caption = "Is Acceleration Envelope >= 1.3 G"  'Mid Userform Location between Label33 and Units jack inputs various information i.e. Turn knob, push button, wait till stable etc

            
        ElseIf TestSect = 6 Then
            'This is the Function Description for example AC Voltage Tests @ 60 Hz - From Datasheet Just click the test description and get the cell address for example B14
            'Then change the address in the quotes below
            Me.Caption = dataSheet.Range("B57").Value
            'this is the path to the hookup image, if using Naming scheme(highly recommended) it will look up the correct image for the unit and Standard model numbers
            imagePath = ThisWorkbook.Path & "\Images\" & Model & "\" & CalibModel & "\E Main Hookup " & CalibModel & ".jpg"
            
            'This is the Title of the userForm at the very top
            Me.Label6.Caption = "Connect to " & CalibModel & " to Test Audio Output"
            'This is the main information for the test ie Turn Knob, push button, eat a snack etc.
            Me.Label33.Caption = "Press On/Select to Select E"
            Me.Label35.Visible = False  '3458 Userform True to show False to not show
            
            Me.Pass.Visible = True
            Me.Fail.Visible = True
            Me.Advance.Visible = False
            
            'Unit to Calibrator/DMM Connections
            Me.Label32.Caption = "BNC"              'Wiring from Meter
            Me.Label31.Caption = "<---->"           '<---->  Hookup between Unit and standards
            Me.Label30.Caption = "Normal Hi"        'Wiring to Calibrator/DMM
            Me.Label29.Caption = "BNC"              'Wiring from Meter
            Me.Label28.Caption = "<---->"           '<---->  Hookup between Unit and standards
            Me.Label27.Caption = "Normal Lo"        'Wiring to Calibrator/DMM
            Me.Label26.Caption = ""                 'Wiring from Meter
            Me.Label25.Caption = ""                 '<---->  Hookup between Unit and standards
            Me.Label24.Caption = ""                 'Wiring to Calibrator/DMM
            Me.Label14.Caption = ""                 'Wiring from Meter
            Me.Label23.Caption = ""                 '<---->  Hookup between Unit and standards
            Me.Label22.Caption = ""                 'Wiring to Calibrator/DMM
            Me.Label34.Caption = "Audio Heard in Headphones?"  'Mid Userform Location between Label33 and Units jack inputs various information i.e. Turn knob, push button, wait till stable etc
            
        ElseIf TestSect = 7 Then
            'This is the Function Description for example AC Voltage Tests @ 60 Hz - From Datasheet Just click the test description and get the cell address for example B14
            'Then change the address in the quotes below
            Me.Caption = dataSheet.Range("B59").Value
            'this is the path to the hookup image, if using Naming scheme(highly recommended) it will look up the correct image for the unit and Standard model numbers
            imagePath = ThisWorkbook.Path & "\Images\" & Model & "\" & CalibModel & "\E Main Hookup " & CalibModel & ".jpg"
            
            'This is the Title of the userForm at the very top
            Me.Label6.Caption = "Connect to " & CalibModel & " to Test Volume Control"
            'This is the main information for the test ie Turn Knob, push button, eat a snack etc.
            Me.Label33.Caption = "Press On/Select to Select E"
            Me.Label35.Visible = False  '3458 Userform True to show False to not show
            
            Me.Pass.Visible = True
            Me.Fail.Visible = True
            Me.Advance.Visible = False
            
            'Unit to Calibrator/DMM Connections
            Me.Label32.Caption = "BNC"              'Wiring from Meter
            Me.Label31.Caption = "<---->"           '<---->  Hookup between Unit and standards
            Me.Label30.Caption = "Normal Hi"        'Wiring to Calibrator/DMM
            Me.Label29.Caption = "BNC"              'Wiring from Meter
            Me.Label28.Caption = "<---->"           '<---->  Hookup between Unit and standards
            Me.Label27.Caption = "Normal Lo"        'Wiring to Calibrator/DMM
            Me.Label26.Caption = ""                 'Wiring from Meter
            Me.Label25.Caption = ""                 '<---->  Hookup between Unit and standards
            Me.Label24.Caption = ""                 'Wiring to Calibrator/DMM
            Me.Label14.Caption = ""                 'Wiring from Meter
            Me.Label23.Caption = ""                 '<---->  Hookup between Unit and standards
            Me.Label22.Caption = ""                 'Wiring to Calibrator/DMM
            Me.Label34.Caption = "Does the Volume Control Function?"  'Mid Userform Location between Label33 and Units jack inputs various information i.e. Turn knob, push button, wait till stable etc
            
        ElseIf TestSect = 8 Then
            'This is the Function Description for example AC Voltage Tests @ 60 Hz - From Datasheet Just click the test description and get the cell address for example B14
            'Then change the address in the quotes below
            Me.Caption = dataSheet.Range("B61").Value
            'this is the path to the hookup image, if using Naming scheme(highly recommended) it will look up the correct image for the unit and Standard model numbers
            'imagePath = ThisWorkbook.Path & "\Images\" & Model & "\" & CalibModel & "\Main Hookup " & CalibModel & ".jpg"
            
            'This is the Title of the userForm at the very top
            Me.Label6.Caption = "Enter Transducer Information, and Verification readings"
            'This is the main information for the test ie Turn Knob, push button, eat a snack etc.
            Me.Label33.Visible = False
            Me.Label35.Visible = False  '3458 Userform True to show False to not show
            
            'Unit to Calibrator/DMM Connections
            Me.Label32.Visible = False   'Wiring from Meter
            Me.Label31.Visible = False           '<---->  Hookup between Unit and standards
            Me.Label30.Visible = False        'Wiring to Calibrator/DMM
            Me.Label29.Visible = False              'Wiring from Meter
            Me.Label28.Visible = False           '<---->  Hookup between Unit and standards
            Me.Label27.Visible = False        'Wiring to Calibrator/DMM
            Me.Label26.Visible = False                 'Wiring from Meter
            Me.Label25.Visible = False                 '<---->  Hookup between Unit and standards
            Me.Label24.Visible = False                 'Wiring to Calibrator/DMM
            Me.Label14.Visible = False                 'Wiring from Meter
            Me.Label23.Visible = False                 '<---->  Hookup between Unit and standards
            Me.Label22.Visible = False                 'Wiring to Calibrator/DMM
            Me.Label34.Visible = False  'Mid Userform Location between Label33 and Units jack inputs various information i.e. Turn knob, push button, wait till stable etc
            Me.Label36.Visible = False
            
        ElseIf TestSect = 9 Then
            'This is the Function Description for example AC Voltage Tests @ 60 Hz - From Datasheet Just click the test description and get the cell address for example B14
            'Then change the address in the quotes below
            Me.Caption = dataSheet.Range("B47").Value
            'this is the path to the hookup image, if using Naming scheme(highly recommended) it will look up the correct image for the unit and Standard model numbers
            imagePath = ThisWorkbook.Path & "\Images\" & Model & "\" & CalibModel & "\Main Hookup " & CalibModel & ".jpg"
            
            'This is the Title of the userForm at the very top
            Me.Label6.Caption = "Connect to " & CalibModel & " to Read Ohms"
            'This is the main information for the test ie Turn Knob, push button, eat a snack etc.
            Me.Label33.Caption = "Turn knob to Ohm-Cont-Diode"
            Me.Label35.Visible = False  '3458 Userform True to show False to not show
            
            'Unit to Calibrator/DMM Connections
            Me.Label32.Caption = "Vohm/Ohm/Diode"   'Wiring from Meter
            Me.Label31.Caption = "<---->"           '<---->  Hookup between Unit and standards
            Me.Label30.Caption = "Normal Hi"        'Wiring to Calibrator/DMM
            Me.Label29.Caption = "Com"              'Wiring from Meter
            Me.Label28.Caption = "<---->"           '<---->  Hookup between Unit and standards
            Me.Label27.Caption = "Normal Lo"        'Wiring to Calibrator/DMM
            Me.Label26.Caption = ""                 'Wiring from Meter
            Me.Label25.Caption = ""                 '<---->  Hookup between Unit and standards
            Me.Label24.Caption = ""                 'Wiring to Calibrator/DMM
            Me.Label14.Caption = ""                 'Wiring from Meter
            Me.Label23.Caption = ""                 '<---->  Hookup between Unit and standards
            Me.Label22.Caption = ""                 'Wiring to Calibrator/DMM
            Me.Label34.Caption = ""  'Mid Userform Location between Label33 and Units jack inputs various information i.e. Turn knob, push button, wait till stable etc
            
        ElseIf TestSect = 10 Then
            'This is the Function Description for example AC Voltage Tests @ 60 Hz - From Datasheet Just click the test description and get the cell address for example B14
            'Then change the address in the quotes below
            Me.Caption = dataSheet.Range("B47").Value
            'this is the path to the hookup image, if using Naming scheme(highly recommended) it will look up the correct image for the unit and Standard model numbers
            imagePath = ThisWorkbook.Path & "\Images\" & Model & "\" & CalibModel & "\Main Hookup " & CalibModel & ".jpg"
            
            'This is the Title of the userForm at the very top
            Me.Label6.Caption = "Connect to " & CalibModel & " to Read Ohms"
            'This is the main information for the test ie Turn Knob, push button, eat a snack etc.
            Me.Label33.Caption = "Turn knob to Ohm-Cont-Diode"
            Me.Label35.Visible = False  '3458 Userform True to show False to not show
            
            'Unit to Calibrator/DMM Connections
            Me.Label32.Caption = "Vohm/Ohm/Diode"   'Wiring from Meter
            Me.Label31.Caption = "<---->"           '<---->  Hookup between Unit and standards
            Me.Label30.Caption = "Normal Hi"        'Wiring to Calibrator/DMM
            Me.Label29.Caption = "Com"              'Wiring from Meter
            Me.Label28.Caption = "<---->"           '<---->  Hookup between Unit and standards
            Me.Label27.Caption = "Normal Lo"        'Wiring to Calibrator/DMM
            Me.Label26.Caption = ""                 'Wiring from Meter
            Me.Label25.Caption = ""                 '<---->  Hookup between Unit and standards
            Me.Label24.Caption = ""                 'Wiring to Calibrator/DMM
            Me.Label14.Caption = ""                 'Wiring from Meter
            Me.Label23.Caption = ""                 '<---->  Hookup between Unit and standards
            Me.Label22.Caption = ""                 'Wiring to Calibrator/DMM
            Me.Label34.Caption = ""  'Mid Userform Location between Label33 and Units jack inputs various information i.e. Turn knob, push button, wait till stable etc
            
        ElseIf TestSect = 11 Then
            'This is the Function Description for example AC Voltage Tests @ 60 Hz - From Datasheet Just click the test description and get the cell address for example B14
            'Then change the address in the quotes below
            Me.Caption = dataSheet.Range("B47").Value
            'this is the path to the hookup image, if using Naming scheme(highly recommended) it will look up the correct image for the unit and Standard model numbers
            imagePath = ThisWorkbook.Path & "\Images\" & Model & "\" & CalibModel & "\Main Hookup " & CalibModel & ".jpg"
            
            'This is the Title of the userForm at the very top
            Me.Label6.Caption = "Connect to " & CalibModel & " to Read Ohms"
            'This is the main information for the test ie Turn Knob, push button, eat a snack etc.
            Me.Label33.Caption = "Turn knob to Ohm-Cont-Diode"
            Me.Label35.Visible = False  '3458 Userform True to show False to not show
            
            'Unit to Calibrator/DMM Connections
            Me.Label32.Caption = "Vohm/Ohm/Diode"   'Wiring from Meter
            Me.Label31.Caption = "<---->"           '<---->  Hookup between Unit and standards
            Me.Label30.Caption = "Normal Hi"        'Wiring to Calibrator/DMM
            Me.Label29.Caption = "Com"              'Wiring from Meter
            Me.Label28.Caption = "<---->"           '<---->  Hookup between Unit and standards
            Me.Label27.Caption = "Normal Lo"        'Wiring to Calibrator/DMM
            Me.Label26.Caption = ""                 'Wiring from Meter
            Me.Label25.Caption = ""                 '<---->  Hookup between Unit and standards
            Me.Label24.Caption = ""                 'Wiring to Calibrator/DMM
            Me.Label14.Caption = ""                 'Wiring from Meter
            Me.Label23.Caption = ""                 '<---->  Hookup between Unit and standards
            Me.Label22.Caption = ""                 'Wiring to Calibrator/DMM
            Me.Label34.Caption = ""  'Mid Userform Location between Label33 and Units jack inputs various information i.e. Turn knob, push button, wait till stable etc
            
        ElseIf TestSect = 12 Then
            'This is the Function Description for example AC Voltage Tests @ 60 Hz - From Datasheet Just click the test description and get the cell address for example B14
            'Then change the address in the quotes below
            Me.Caption = dataSheet.Range("B65").Value
            'this is the path to the hookup image, if using Naming scheme(highly recommended) it will look up the correct image for the unit and Standard model numbers
            imagePath = ThisWorkbook.Path & "\Images\" & Model & "\" & CalibModel & "\Main Hookup " & CalibModel & ".jpg"
            
            'This is the Title of the userForm at the very top
            Me.Label6.Caption = "Connect to " & CalibModel & " to Test Continuity"
            'This is the main information for the test ie Turn Knob, push button, eat a snack etc.
            Me.Label33.Caption = "Turn knob to Ohm-Cont-Diode"
            Me.Label35.Visible = False
            
            'Unit to Calibrator/DMM Connections
            Me.Label32.Caption = "Vohm/Ohm/Diode"   'Wiring from Meter
            Me.Label31.Caption = "<---->"           '<---->  Hookup between Unit and standards
            Me.Label30.Caption = "Normal Hi"        'Wiring to Calibrator/DMM
            Me.Label29.Caption = "Com"              'Wiring from Meter
            Me.Label28.Caption = "<---->"           '<---->  Hookup between Unit and standards
            Me.Label27.Caption = "Normal Lo"        'Wiring to Calibrator/DMM
            Me.Label26.Caption = ""                 'Wiring from Meter
            Me.Label25.Caption = ""                 '<---->  Hookup between Unit and standards
            Me.Label24.Caption = ""                 'Wiring to Calibrator/DMM
            Me.Label14.Caption = ""                 'Wiring from Meter
            Me.Label23.Caption = ""                 '<---->  Hookup between Unit and standards
            Me.Label22.Caption = ""                 'Wiring to Calibrator/DMM
            Me.Label34.Caption = "Press (((( button to activate Continuity"  'Mid Userform Location between Label33 and Units jack inputs various information i.e. Turn knob, push button, wait till stable etc
            
        ElseIf TestSect = 13 Then
            'This is the Function Description for example AC Voltage Tests @ 60 Hz - From Datasheet Just click the test description and get the cell address for example B14
            'Then change the address in the quotes below
            Me.Caption = dataSheet.Range("B66").Value
            'this is the path to the hookup image, if using Naming scheme(highly recommended) it will look up the correct image for the unit and Standard model numbers
            imagePath = ThisWorkbook.Path & "\Images\" & Model & "\" & CalibModel & "\Main Hookup " & CalibModel & ".jpg"
            
            'This is the Title of the userForm at the very top
            Me.Label6.Caption = "Connect to " & CalibModel & " to Test Continuity"
            'This is the main information for the test ie Turn Knob, push button, eat a snack etc.
            Me.Label33.Caption = "Turn knob to Ohm-Cont-Diode"
            Me.Label35.Visible = False
            
            'Unit to Calibrator/DMM Connections
            Me.Label32.Caption = "Vohm/Ohm/Diode"   'Wiring from Meter
            Me.Label31.Caption = "<---->"           '<---->  Hookup between Unit and standards
            Me.Label30.Caption = "Normal Hi"        'Wiring to Calibrator/DMM
            Me.Label29.Caption = "Com"              'Wiring from Meter
            Me.Label28.Caption = "<---->"           '<---->  Hookup between Unit and standards
            Me.Label27.Caption = "Normal Lo"        'Wiring to Calibrator/DMM
            Me.Label26.Caption = ""                 'Wiring from Meter
            Me.Label25.Caption = ""                 '<---->  Hookup between Unit and standards
            Me.Label24.Caption = ""                 'Wiring to Calibrator/DMM
            Me.Label14.Caption = ""                 'Wiring from Meter
            Me.Label23.Caption = ""                 '<---->  Hookup between Unit and standards
            Me.Label22.Caption = ""                 'Wiring to Calibrator/DMM
            Me.Label34.Caption = "Press (((( button to activate Continuity"  'Mid Userform Location between Label33 and Units jack inputs various information i.e. Turn knob, push button, wait till stable etc
            
        ElseIf TestSect = 14 Then
            'This is the Function Description for example AC Voltage Tests @ 60 Hz - From Datasheet Just click the test description and get the cell address for example B14
            'Then change the address in the quotes below
            Me.Caption = dataSheet.Range("B68").Value
            'this is the path to the hookup image, if using Naming scheme(highly recommended) it will look up the correct image for the unit and Standard model numbers
            imagePath = ThisWorkbook.Path & "\Images\" & Model & "\" & CalibModel & "\Main Hookup " & CalibModel & ".jpg"
            
            'This is the Title of the userForm at the very top
            Me.Label6.Caption = "Connect to " & CalibModel & " to Test Diode"
            'This is the main information for the test ie Turn Knob, push button, eat a snack etc.
            Me.Label33.Caption = "Turn knob to Ohm-Cont-Diode"
            Me.Label35.Visible = False  '3458 Userform True to show False to not show
            
            'Unit to Calibrator/DMM Connections
            Me.Label32.Caption = "Vohm/Ohm/Diode"   'Wiring from Meter
            Me.Label31.Caption = "<---->"           '<---->  Hookup between Unit and standards
            Me.Label30.Caption = "Normal Hi"        'Wiring to Calibrator/DMM
            Me.Label29.Caption = "Com"              'Wiring from Meter
            Me.Label28.Caption = "<---->"           '<---->  Hookup between Unit and standards
            Me.Label27.Caption = "Normal Lo"        'Wiring to Calibrator/DMM
            Me.Label26.Caption = ""                 'Wiring from Meter
            Me.Label25.Caption = ""                 '<---->  Hookup between Unit and standards
            Me.Label24.Caption = ""                 'Wiring to Calibrator/DMM
            Me.Label14.Caption = ""                 'Wiring from Meter
            Me.Label23.Caption = ""                 '<---->  Hookup between Unit and standards
            Me.Label22.Caption = ""                 'Wiring to Calibrator/DMM
            Me.Label34.Caption = "Press Blue button to select Diode"  'Mid Userform Location between Label33 and Units jack inputs various information i.e. Turn knob, push button, wait till stable etc
            
        ElseIf TestSect = 15 Then
            'This is the Function Description for example AC Voltage Tests @ 60 Hz - From Datasheet Just click the test description and get the cell address for example B14
            'Then change the address in the quotes below
            Me.Caption = dataSheet.Range("B68").Value
            'this is the path to the hookup image, if using Naming scheme(highly recommended) it will look up the correct image for the unit and Standard model numbers
            imagePath = ThisWorkbook.Path & "\Images\" & Model & "\" & DMMModel & "\Diode " & DMMModel & ".jpg"
            
            'This is the Title of the userForm at the very top
            Me.Label6.Caption = "Connect to " & DMMModel & " to Test Diode"
            
            'This is the main information for the test ie Turn Knob, push button, eat a snack etc.
            'Comment out one that is not being used below
            'Calibrator Image
            'Me.Label33.Caption = "Turn knob to Ohm-Cont-Diode"
            Me.Label33.Visible = False
            'DMM Image
            
            
            'Put a 1 for Calibrator or 2 for DMM, that is being used for this test in the STDStyle
            STDStyle = 2
            'Fill out the Captions
            If STDStyle = 1 Then
                Me.Label33.Caption = "Turn knob to Ohm-Cont-Diode"
                Me.Label34.Caption = "test"  'Mid Userform Location between Label33 and Units jack inputs various information i.e. Turn knob, push button, wait till stable etc
                Me.Label33.Visible = True
                Me.Label34.Visible = True
            ElseIf STDStyle = 2 Then
                Me.Label35.Caption = "Turn knob to Ohm-Cont-Diode. Press the Blue Button."
                Me.Label34.Caption = ""  'Mid Userform Location between Label33 and Units jack inputs various information i.e. Turn knob, push button, wait till stable etc
                Me.Label34.Visible = False
                Me.Label33.Visible = False
            End If
            
            
            'Unit to Calibrator/DMM Connections
            Me.Label32.Caption = "Vohm/Ohm/Diode"   'Wiring from Meter
            Me.Label31.Caption = "<---->"           '<---->  Hookup between Unit and standards
            Me.Label30.Caption = "Input I"        'Wiring to Calibrator/DMM
            Me.Label29.Caption = "Com"              'Wiring from Meter
            Me.Label28.Caption = "<---->"           '<---->  Hookup between Unit and standards
            Me.Label27.Caption = "Input Lo"        'Wiring to Calibrator/DMM
            Me.Label26.Caption = ""                 'Wiring from Meter
            Me.Label25.Caption = ""                 '<---->  Hookup between Unit and standards
            Me.Label24.Caption = ""                 'Wiring to Calibrator/DMM
            Me.Label14.Caption = ""                 'Wiring from Meter
            Me.Label23.Caption = ""                 '<---->  Hookup between Unit and standards
            Me.Label22.Caption = ""                 'Wiring to Calibrator/DMM
            Me.Label34.Visible = False  'Mid Userform Location between Label33 and Units jack inputs various information i.e. Turn knob, push button, wait till stable etc
            
        ElseIf TestSect = 16 Then
            'This is the Function Description for example AC Voltage Tests @ 60 Hz - From Datasheet Just click the test description and get the cell address for example B14
            'Then change the address in the quotes below
            Me.Caption = dataSheet.Range("B72").Value
            'this is the path to the hookup image, if using Naming scheme(highly recommended) it will look up the correct image for the unit and Standard model numbers
            imagePath = ThisWorkbook.Path & "\Images\" & Model & "\" & CalibModel & "\DC mA " & CalibModel & ".jpg"
            
            'This is the Title of the userForm at the very top
            Me.Label6.Caption = "Connect to " & CalibModel & " to Read mA"
            
            'This is the main information for the test ie Turn Knob, push button, eat a snack etc.
            'Put a 1 for Calibrator or 2 for DMM, that is being used for this test in the STDStyle
            'Then fill out the captions
            STDStyle = 1
            'Fill out the Captions
            If STDStyle = 1 Then
                Me.Label33.Caption = "Turn knob to mA/A" 'This is the main test description you can add a couple of short lines for the test.
                Me.Label34.Caption = ""  'Mid Userform Location between Label33 and Units jack inputs various information i.e. Turn knob, push button, wait till stable etc
                
                'These below just enable or disable some of the text boxes on the userform do not edit -unless you need to.
                Me.Label33.Visible = True
                Me.Label34.Visible = True
                Me.Label35.Visible = False
                
            ElseIf STDStyle = 2 Then
                Me.Label35.Caption = "" 'This is the main test description you can add a couple of short lines for the test.
                Me.Label34.Caption = ""
                
                'These below just enable or disable some of the text boxes on the userform do not edit -unless you need to.
                Me.Label34.Visible = False
                Me.Label33.Visible = False
            End If
            
            
            'Unit to Calibrator/DMM Connections
            Me.Label32.Caption = "mA"   'Wiring from Meter
            Me.Label31.Caption = "<---->"           '<---->  Hookup between Unit and standards
            Me.Label30.Caption = "Aux Hi"        'Wiring to Calibrator/DMM
            Me.Label29.Caption = "Com"              'Wiring from Meter
            Me.Label28.Caption = "<---->"           '<---->  Hookup between Unit and standards
            Me.Label27.Caption = "Aux Lo"        'Wiring to Calibrator/DMM
            Me.Label26.Caption = ""                 'Wiring from Meter
            Me.Label25.Caption = ""                 '<---->  Hookup between Unit and standards
            Me.Label24.Caption = ""                 'Wiring to Calibrator/DMM
            Me.Label14.Caption = ""                 'Wiring from Meter
            Me.Label23.Caption = ""                 '<---->  Hookup between Unit and standards
            Me.Label22.Caption = ""                 'Wiring to Calibrator/DMM
            
            
        ElseIf TestSect = 17 Then
            'This is the Function Description for example AC Voltage Tests @ 60 Hz - From Datasheet Just click the test description and get the cell address for example B14
            'Then change the address in the quotes below
            Me.Caption = dataSheet.Range("B77").Value
            'this is the path to the hookup image, if using Naming scheme(highly recommended) it will look up the correct image for the unit and Standard model numbers
            imagePath = ThisWorkbook.Path & "\Images\" & Model & "\" & CalibModel & "\DC Amps " & CalibModel & ".jpg"
            
            'This is the Title of the userForm at the very top
            Me.Label6.Caption = "Connect to " & CalibModel & " to Read DC A"
            
            'This is the main information for the test ie Turn Knob, push button, eat a snack etc.
            'Put a 1 for Calibrator or 2 for DMM, that is being used for this test in the STDStyle
            'Then fill out the captions
            STDStyle = 1
            'Fill out the Captions
            If STDStyle = 1 Then
                Me.Label33.Caption = "Turn knob to mA/A" 'This is the main test description you can add a couple of short lines for the test.
                Me.Label34.Caption = "Move lead from 30mA to 1A input"  'Mid Userform Location between Label33 and Units jack inputs various information i.e. Turn knob, push button, wait till stable etc
                
                'These below just enable or disable some of the text boxes on the userform do not edit -unless you need to.
                Me.Label33.Visible = True
                Me.Label34.Visible = True
                Me.Label35.Visible = False
                
            ElseIf STDStyle = 2 Then
                Me.Label35.Caption = "" 'This is the main test description you can add a couple of short lines for the test.
                Me.Label34.Caption = ""
                
                'These below just enable or disable some of the text boxes on the userform do not edit -unless you need to.
                Me.Label34.Visible = False
                Me.Label33.Visible = False
            End If
            
            
            'Unit to Calibrator/DMM Connections
            Me.Label32.Caption = "A AC/DC"            'Wiring from Meter
            Me.Label31.Caption = "<---->"           '<---->  Hookup between Unit and standards
            Me.Label30.Caption = "Aux Hi"           'Wiring to Calibrator/DMM
            Me.Label29.Caption = "Com"              'Wiring from Meter
            Me.Label28.Caption = "<---->"           '<---->  Hookup between Unit and standards
            Me.Label27.Caption = "Aux Lo"           'Wiring to Calibrator/DMM
            Me.Label26.Caption = ""                 'Wiring from Meter
            Me.Label25.Caption = ""                 '<---->  Hookup between Unit and standards
            Me.Label24.Caption = ""                 'Wiring to Calibrator/DMM
            Me.Label14.Caption = ""                 'Wiring from Meter
            Me.Label23.Caption = ""                 '<---->  Hookup between Unit and standards
            Me.Label22.Caption = ""                 'Wiring to Calibrator/DMM
            
        ElseIf TestSect = 18 Then
            'This is the Function Description for example AC Voltage Tests @ 60 Hz - From Datasheet Just click the test description and get the cell address for example B14
            'Then change the address in the quotes below
            Me.Caption = dataSheet.Range("B81").Value
            'this is the path to the hookup image, if using Naming scheme(highly recommended) it will look up the correct image for the unit and Standard model numbers
            imagePath = ThisWorkbook.Path & "\Images\" & Model & "\" & CalibModel & "\AC Amps " & CalibModel & ".jpg"
            
            'This is the Title of the userForm at the very top
            Me.Label6.Caption = "Connect to " & CalibModel & " to Read AC A"
            
            'This is the main information for the test ie Turn Knob, push button, eat a snack etc.
            'Put a 1 for Calibrator or 2 for DMM, that is being used for this test in the STDStyle
            'Then fill out the captions
            STDStyle = 1
            'Fill out the Captions
            If STDStyle = 1 Then
                Me.Label33.Caption = "Turn knob to mA/A" 'This is the main test description you can add a couple of short lines for the test.
                Me.Label34.Caption = "Press Blue Button to switch to AC Amps"  'Mid Userform Location between Label33 and Units jack inputs various information i.e. Turn knob, push button, wait till stable etc
                
                'These below just enable or disable some of the text boxes on the userform do not edit -unless you need to.
                Me.Label33.Visible = True
                Me.Label34.Visible = True
                Me.Label35.Visible = False
                'Label13 is for the standard Model being used. Just comment out the one not being used.
                Me.Label13.Caption = CalibModel
                
            ElseIf STDStyle = 2 Then
                Me.Label35.Caption = "" 'This is the main test description you can add a couple of short lines for the test.
                Me.Label34.Caption = ""
                
                'These below just enable or disable some of the text boxes on the userform do not edit -unless you need to.
                Me.Label34.Visible = False
                Me.Label33.Visible = False
                Me.Label13.Caption = DMModel
            End If
            
   
            
            'Unit to Calibrator/DMM Connections
            Me.Label32.Caption = "1 Amp"            'Wiring from Meter
            Me.Label31.Caption = "<---->"           '<---->  Hookup between Unit and standards
            Me.Label30.Caption = "Aux Hi"           'Wiring to Calibrator/DMM
            Me.Label29.Caption = "Com"              'Wiring from Meter
            Me.Label28.Caption = "<---->"           '<---->  Hookup between Unit and standards
            Me.Label27.Caption = "Aux Lo"           'Wiring to Calibrator/DMM
            Me.Label26.Caption = ""                 'Wiring from Meter
            Me.Label25.Caption = ""                 '<---->  Hookup between Unit and standards
            Me.Label24.Caption = ""                 'Wiring to Calibrator/DMM
            Me.Label14.Caption = ""                 'Wiring from Meter
            Me.Label23.Caption = ""                 '<---->  Hookup between Unit and standards
            Me.Label22.Caption = ""                 'Wiring to Calibrator/DMM
               
        ElseIf TestSect = 19 Then
            'This is the Function Description for example AC Voltage Tests @ 60 Hz - From Datasheet Just click the test description and get the cell address for example B14
            'Then change the address in the quotes below
            Me.Caption = dataSheet.Range("B85").Value
            'this is the path to the hookup image, if using Naming scheme(highly recommended) it will look up the correct image for the unit and Standard model numbers
            imagePath = ThisWorkbook.Path & "\Images\" & Model & "\" & DMMModel & "\DC mA Source " & DMMModel & ".jpg"
            
            'This is the Title of the userForm at the very top
            Me.Label6.Caption = "Connect to " & DMMModel & " to Source mA"
            
            'This is the main information for the test ie Turn Knob, push button, eat a snack etc.
            'Put a 1 for Calibrator or 2 for DMM, that is being used for this test in the STDStyle
            'Then fill out the captions
            STDStyle = 2
            'Fill out the Captions
            If STDStyle = 1 Then
                Me.Label33.Caption = "Turn knob to mA/A" 'This is the main test description you can add a couple of short lines for the test.
                Me.Label34.Caption = "Press Blue Button to switch to AC Amps"  'Mid Userform Location between Label33 and Units jack inputs various information i.e. Turn knob, push button, wait till stable etc
                
                'These below just enable or disable some of the text boxes on the userform do not edit -unless you need to.
                Me.Label33.Visible = True
                Me.Label34.Visible = True
                Me.Label35.Visible = False
                'Label13 is for the standard Model being used. Just comment out the one not being used.
                Me.Label13.Caption = CalibModel
                
            ElseIf STDStyle = 2 Then
                Me.Label35.Caption = "Turn Knob to mA Output" 'This is the main test description you can add a couple of short lines for the test.
                Me.Label34.Caption = ""
                
                'These below just enable or disable some of the text boxes on the userform do not edit -unless you need to.
                Me.Label34.Visible = False
                Me.Label33.Visible = False
                Me.Label13.Caption = DMMModel
            End If
            
   
            
            'Unit to Calibrator/DMM Connections
            Me.Label32.Caption = "Source +"            'Wiring from Meter
            Me.Label31.Caption = "<---->"           '<---->  Hookup between Unit and standards
            Me.Label30.Caption = "Normal I"           'Wiring to Calibrator/DMM
            Me.Label29.Caption = "Source -"              'Wiring from Meter
            Me.Label28.Caption = "<---->"           '<---->  Hookup between Unit and standards
            Me.Label27.Caption = "Normal Lo"           'Wiring to Calibrator/DMM
            Me.Label26.Caption = ""                 'Wiring from Meter
            Me.Label25.Caption = ""                 '<---->  Hookup between Unit and standards
            Me.Label24.Caption = ""                 'Wiring to Calibrator/DMM
            Me.Label14.Caption = ""                 'Wiring from Meter
            Me.Label23.Caption = ""                 '<---->  Hookup between Unit and standards
            Me.Label22.Caption = ""                 'Wiring to Calibrator/DMM
           
        ElseIf TestSect = 20 Then
            
        
        ElseIf TestSect = 21 Then
            
            
        ElseIf TestSect = 22 Then
            
            
        ElseIf TestSect = 23 Then
            
                
        ElseIf TestSect = 24 Then
            
                    
        ElseIf TestSect = 25 Then
            
                        
        ElseIf TestSect = 26 Then
            
                            
        ElseIf TestSect = 27 Then
            
                            
        ElseIf TestSect = 28 Then
            
                            
        ElseIf TestSect = 29 Then
            
                            
        ElseIf TestSect = 30 Then
            
            
            
        End If

            
    
    
    
    If Dir(imagePath) <> "" Then
        Me.Image1.Picture = LoadPicture(imagePath)
    Else
        MsgBox "Image not found: " & imagePath, vbExclamation
    End If

End Sub


