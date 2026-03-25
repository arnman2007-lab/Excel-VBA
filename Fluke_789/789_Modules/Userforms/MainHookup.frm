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
Public Title As String
Public Comments As String
Public AdditComments As String
Public SectTitle As String
Public ImageName As String
Public imagePath As String

Private Sub Advance_Click()

shouldContinue = True
If HRSTextInput <> "" Then
ActiveCell.Value = HRSTextInput.Value
End If
Unload Me
End Sub

Private Sub Fail_Click()
ActiveCell.Value = "Fail"
Unload Me
Selection.OffSet(1, 0).Select
End Sub

Private Sub Image2_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal DragState As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)

End Sub





Private Sub Pass_Click()
ActiveCell.Value = "Pass"
Unload Me
Selection.OffSet(1, 0).Select
End Sub

Private Sub Terminate_Click()
    TerminateClicked = True
    ActiveCell.OffSet(0, -2).Select
    PrevAddress = ActiveCell.Address
    shouldContinue = False

    
Unload Me
    
End Sub
Private Sub TextInput_Enter()
    'When the user clicks inside
    If TextInput.ForeColor = vbGrayText Then
        TextInput.Text = ""
        TextInput.ForeColor = vbBlack
    End If
End Sub

Private Sub TextInput_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    'When the user leaves the box
    If Trim(TextInput.Text) = "" Then
        TextInput.Text = "Enter Reading"
        TextInput.ForeColor = vbGrayText
    End If
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
   'Set initial placeholder
    TextInput.Text = "Enter Reading"
    TextInput.ForeColor = vbGrayText
   
    
    
        'Copy and paste the insides, (between If TestSect = 1 Then and ElseIf TestSect = 2 Then - just the code), of TestSect1 into the ElseIf TestSect = 2 Then Fill out
        
        If TestSect = 1 Then
            'This is the Function Description for example AC Voltage Tests @ 60 Hz - From Datasheet Just click the test description and get the cell address for example B14
            'Then change the address in the quotes below example C12, B4
            SectTitle = "B12"
            
            
            'This is the Image name do not use the full image you created name in the quotes below
            'If your image name is mAUpper Main Hookup 5520A
            'Just use mAUpper Main Hookup  with no spaces in the front or back
            'like this  ImageName = "mAUpper Main Hookup" leave the STD model out the code will do the rest
            ImageName = "Main Hookup"
            
            
            'This is the Title of the userForm at the very top This could be anything informative
            'You do not have to add a CalibratorModel ie 5502A variable just anything between quotes
            'Example
            'Title = "This is a Test of the National Broadcast System" you have some many letters so chose wisely
            Title = "Connect " & Model & " to " & CalibratorModel
            
            
            'This is the main information for the test ie Turn Knob, push button, eat a snack etc.
            Comments = "Turn knob to ACV to measure Voltage"
            
            
            'Additional Comments between Label33 and Units jack inputs various information i.e. Turn knob, push button, wait till stable etc
            'This only works with the calibrator setup image not DMM
            AdditComments = ""

            
            
            'Setup your wiring hookups here Just type your input output jack in the quotes after telling it is a DMM or Calibrator
            '         Calibrator/DMM UnitJack1     STDJack1 UnitJack2  STDJack2   UnitJack3  STDJack3  UnitJack4  STDJack4
            'Hookup    "Calibrator", "",             "",     "Com",  "Normal Lo", "",          "",       "",        ""
            'Example
            'Hookup "DMM or Calibrator", "", "", "", "", "", "", "", ""
            'Hookup "Calibrator", "V mA Loop", "Normal Hi", "Com", "Normal Lo", "", "", "", ""
            Hookup "Calibrator", "V/Ohm/Diode", "Normal Hi", "Com", "Normal Lo", "", "", "", ""
            

                        
           
            
        ElseIf TestSect = 2 Then
            
            ImageName = "Main Hookup"
            
            SectTitle = "B12"
            
            Title = "Connect " & Model & " to " & CalibratorModel
            
            Comments = "Turn Knob to ACV. Press the Range Button to switch 1000V Range."
            
            AdditComments = ""
            
            Hookup "Calibrator", "V/Ohm/Diode", "Normal Hi", "Com", "Normal Lo", "", "", "", ""
            
            
        ElseIf TestSect = 3 Then
        
            ImageName = "Main Hookup"
            
            SectTitle = "B26"
            
            Title = "Connect " & Model & " to " & CalibratorModel
            
            Comments = "Turn Knob to ACV. Press the Hz button to measure hertz."
            
            AdditComments = ""

            Hookup "Calibrator", "V/Ohm/Diode", "Normal Hi", "Com", "Normal Lo", "", "", "", ""
            
            
        ElseIf TestSect = 4 Then
        
            ImageName = "Main Hookup"
            
            SectTitle = "B31"
            
            Title = "Connect " & Model & " to " & CalibratorModel
            
            Comments = "Turn Knob to DCV to measure DC Volts"
            
            AdditComments = ""

            Hookup "Calibrator", "V/Ohm/Diode", "Normal Hi", "Com", "Normal Lo", "", "", "", ""


            
        ElseIf TestSect = 5 Then
        
            ImageName = "Main Hookup"
            
            SectTitle = "B31"
            
            Title = "Connect " & Model & " to " & CalibratorModel
            
            Comments = "Turn Knob to DCV. Press Range button to switch to 1000V Range."
            
            AdditComments = ""

            Hookup "Calibrator", "V/Ohm/Diode", "Normal Hi", "Com", "Normal Lo", "", "", "", ""
            

            
            
        ElseIf TestSect = 6 Then
        
            ImageName = "Main Hookup"
            
            SectTitle = "B40"
            
            Title = "Connect " & Model & " to " & CalibratorModel
            
            Comments = "Turn knob to mV DC to measure DC millivolts."
            
            AdditComments = ""

            Hookup "Calibrator", "V/Ohm/Diode", "Normal Hi", "Com", "Normal Lo", "", "", "", ""

            
            
        ElseIf TestSect = 7 Then
        
            ImageName = "4Wire Ohms"
            
            SectTitle = "B44"
            
            Title = "Connect " & Model & " to " & CalibratorModel
            
            Comments = "Turn knob to ohm/Cont/Diode, to measure resistance."
            
            AdditComments = ""

            Hookup "Calibrator", "V/Ohm/Diode", "Normal Hi", "Com", "Normal Lo", "V/Ohm/Diode", "Aux Hi", "Com", "Aux Lo"


            
        ElseIf TestSect = 8 Then
        
            ImageName = "Main Hookup"
            
            SectTitle = DiffTitle
            
            Title = "Connect " & Model & " to " & CalibratorModel
            
            Comments = "Turn knob to Ohms/Cont/Diode, and press ))))) Button to test Continuity."
            
            AdditComments = ""
            
            Hookup "Calibrator", "V/Ohm/Diode", "Normal Hi", "Com", "", "", "", "", ""
            
            
        ElseIf TestSect = 9 Then
        
            ImageName = "Main Hookup"
            
            SectTitle = "B65"
            
            Title = "Connect " & Model & " to " & CalibratorModel
            
            Comments = "Turn knob to Ohms/Cont/Diode, and press blue button for diode."
            
            AdditComments = ""
            
            Hookup "Calibrator", "V/Ohm/Diode", "Normal Hi", "Com", "Normal Lo", "", "", "", ""
            
            
        ElseIf TestSect = 10 Then
        
            ImageName = "Diode"
            
            SectTitle = "C67"
            
            Title = "Connect " & Model & " to " & CalibratorModel
            
            Comments = "Turn knob to Ohm/Cont/Diode, and press blue button to test diode."
            
            AdditComments = "Unit outputs 0.2-0.33mA?"
            
            
            If DMMGPIB = "" Then
                Hookup "OpCheck", "V/Ohm/Diode", "Input I", "Com", "Input Lo", "", "", "", ""
            Else
                Hookup "DMM", "V/Ohm/Diode", "Input I", "Com", "Input Lo", "", "", "", ""
            End If
            
        ElseIf TestSect = 11 Then
        
            ImageName = "DC mA"
            
            SectTitle = "B69"
            
            Title = "Connect " & Model & " to " & CalibratorModel
            
            Comments = "Turn Knob to mA/A to measure DC mA"
            
            AdditComments = ""
            
            Hookup "Calibrator", "mA", "Aux Hi", "Com", "Aux Lo", "", "", "", ""

            
            
        ElseIf TestSect = 12 Then
        
            ImageName = "DC Amps"
            
            SectTitle = "B74"
            
            Title = "Connect " & Model & " to " & CalibratorModel
            
            Comments = "Turn knob to mA/A to measure DC A."
            
            AdditComments = ""

            Hookup "Calibrator", "A", "Input Hi", "Com", "Input Lo", "", "", "", ""
            
        ElseIf TestSect = 13 Then
        
            ImageName = "DC Amps"
            
            SectTitle = "B78"
            
            Title = "Connect " & Model & " to " & CalibratorModel
            
            Comments = "Turn knob to mA/A, and press the blue button to measure AC A."
            
            AdditComments = ""

            Hookup "Calibrator", "A", "Input Hi", "Com", "Input Lo", "", "", "", ""

            
        ElseIf TestSect = 14 Then
        
            ImageName = "DC mA Source"
            
            SectTitle = "B82"
            
            Title = "Connect UUT to " & DMMModel
            
            Comments = "Turn knob to mA output, to source DC mA"
            
            AdditComments = ""

            Hookup "DMM", "Source +", "Input I", "Source -", "Input Lo", "", "", "", ""
            
            
        ElseIf TestSect = 15 Then
        
            ImageName = "Open Circuit Voltage"
            
            SectTitle = "B88"
            
            Title = "Connect UUT to " & DMMModel
            
            Comments = "Turn knob to Loop Power to source DC Voltage into DMM"
            
            AdditComments = ""

            Hookup "DMM", "Source +", "Input Hi", "Source -", "Input Lo", "", "", "", ""
            
            
        ElseIf TestSect = 16 Then
        
            ImageName = "Open Circuit Voltage"
            
            SectTitle = "B89"
            
            Title = "Connect UUT to " & DMMModel
            
            Comments = "Turn knob to Loop Power and press Blue Button to source DC Voltage into DMM"
            
            AdditComments = ""

            Hookup "DMM", "Source +", "Input Hi", "Source -", "Input Lo", "", "", "", ""
            
            
            
        ElseIf TestSect = 17 Then
        
            ImageName = "250 ohm Hart"
            
            SectTitle = "B90"
            
            Title = "Connect UUT to " & DMMModel
            
            Comments = "Turn knob to Loop Power turn Hart Off, put 1 kohm Resistor across Input on DMM"
            
            AdditComments = ""

            Hookup "DMM", "Source +", "Input Hi", "Source -", "Input Lo", "1 kOhm Resistor", "Input Hi/Input Lo", "", ""
            
            
        ElseIf TestSect = 18 Then
        
            ImageName = "DC mA Source"
            
            SectTitle = "B91"
            
            Title = "Connect UUT to " & DMMModel
            
            Comments = "Turn knob to Loop Power and connect UUT to DMM to source full current"
            
            AdditComments = ""

            Hookup "DMM", "Source +", "Input I", "Source -", "Input Lo", "", "", "", ""
               
        ElseIf TestSect = 19 Then
            
           
        ElseIf TestSect = 20 Then
            
        
        ElseIf TestSect = 21 Then
            
            
        ElseIf TestSect = 22 Then
        
            ImageName = ""
            
            SectTitle = "B112"
            
            Title = "Check Backlight Function"
            
            Comments = "Press the backlight button"
            
            AdditComments = "Does the light come on?"

            Hookup "OpCheck", "", "", "", "", "", "", "", ""
            
            
        ElseIf TestSect = 23 Then
        
            ImageName = ""
            
            SectTitle = "B113"
            
            Title = "Check Battery Function"
            
            Comments = "Press Hold and turn knob to Insulation"
            
            AdditComments = "Is the battery >5.2?"

            Hookup "OpCheck", "", "", "", "", "", "", "", ""
            
                
        ElseIf TestSect = 24 Then
        
            ImageName = ""
            
            SectTitle = "B114"
            
            Title = "Check all Keypad Functions"
            
            Comments = "Press every key on keypad."
            
            AdditComments = "Do they function and beep?"

            Hookup "OpCheck", "", "", "", "", "", "", "", ""
            
                    
        ElseIf TestSect = 25 Then
            
                        
        ElseIf TestSect = 26 Then
            
                            
        ElseIf TestSect = 27 Then
            
                            
        ElseIf TestSect = 28 Then
            
                            
        ElseIf TestSect = 29 Then
            
                            
        ElseIf TestSect = 30 Then
        
        
        ElseIf TestSect = 1000 Then
'--------------------------------------Begin Temp Measure Stabilize--------------------------------
        
            ImageName = ""
            
            SectTitle = "B1000"
            
            Title = "Please Wait While Temperature Stabilizes"
            
            Comments = "Press Advance when Temperature is Stabilized in the Calibrator."
            
            AdditComments = ""

            Hookup "Stabilize", "", "", "", "", "", "", "", ""
'--------------------------------------End Temp Measure Stabilize----------------------------------
        
        
        ElseIf TestSect = 2000 Then
'--------------------------------------Begin Frequency Source From Unit----------------------------
            
            ImageName = ""
            
            SectTitle = "B2000"
            
            Title = "Sourcing Frequency to DMM"
            
            Comments = "Source " & VariableString & " to " & DMMModel
            
            AdditComments = "Press Advance when Stable."

            Hookup "Sourcing", "", "", "", "", "", "", "", ""
'--------------------------------------Begin Frequency Source From Unit----------------------------------
        
        
        ElseIf TestSect = 3000 Then
'--------------------------------------Begin DC mV Source From Unit----------------------------
            
            ImageName = ""
            
            SectTitle = "B3000"
            
            Title = "Sourcing MilliVolts to DMM"
            
            Comments = "Source " & VariableString & " to " & DMMModel
            
            AdditComments = "Press Advance when Stable."

            Hookup "Sourcing", "", "", "", "", "", "", "", ""
'--------------------------------------Begin DC mV Source From Unit----------------------------------
        
        
        ElseIf TestSect = 4000 Then
'--------------------------------------Begin DC V Source From Unit----------------------------
            
            ImageName = ""
            
            SectTitle = "B4000"
            
            Title = "Sourcing DC Volts to DMM"
            
            Comments = "Source " & VariableString & " to " & DMMModel
            
            AdditComments = "Press Advance when Stable."

            Hookup "Sourcing", "", "", "", "", "", "", "", ""
'--------------------------------------Begin DC V Source From Unit----------------------------------
        
        
        ElseIf TestSect = 5000 Then
'--------------------------------------Begin Ohms Source From Unit----------------------------
            
            ImageName = ""
            
            SectTitle = "B5000"
            
            Title = "Sourcing Ohms to DMM"
            
            Comments = "Source " & VariableString & " to " & DMMModel
            
            AdditComments = "Press Advance when Stable."

            Hookup "Sourcing", "", "", "", "", "", "", "", ""
'--------------------------------------Begin Ohms Source From Unit----------------------------------
        
        
        ElseIf TestSect = 6000 Then
'--------------------------------------Begin mA Source From Unit----------------------------
            
            ImageName = ""
            
            SectTitle = "B6000"
            
            Title = "Sourcing Ohms to DMM"
            
            Comments = "Source " & VariableString & " to " & DMMModel
            
            AdditComments = "Press Advance when Stable."

            Hookup "Sourcing", "", "", "", "", "", "", "", ""
'--------------------------------------Begin mA Source From Unit----------------------------------
        
        
        ElseIf TestSect = 7000 Then
'--------------------------------------Begin Insulation Tests HRS Box-----------------------------
            
            ImageName = NewHRS & " Main Hookup"
            
            SectTitle = "B7000"
            
            Title = "Insulation Resistance Tests"
            
            Comments = "Turn knob to Source " & VariableString & " to HRS Resistance Box. Press and Hold Test. Type Reading below when stable."
            
            AdditComments = ""

            Hookup "HRS", "", "", "", "", "", "", "", ""
'--------------------------------------Begin Insulation Tests HRS Box----------------------------------
        
        
        ElseIf TestSect = 8000 Then
'--------------------------------------Begin Button Press example Low Pass Filter On-----------------------------
            
            ImageName = ""
            
            SectTitle = "B19"
            
            Title = "Turn " & VariableString & " Low Pass Filter"
            
            Comments = "Press Blue Button to turn " & VariableString & " Low Pass Filter"
            
            AdditComments = ""

            Hookup "ButtonAction", "", "", "", "", "", "", "", ""
'--------------------------------------End Button Press example Low Pass Filter On-----------------------------
        
        
        ElseIf TestSect = 8001 Then
'--------------------------------------Begin Button Press example Low Pass Filter Off-----------------------------
            
            ImageName = ""
            
            SectTitle = "B21"
            
            Title = "Turn " & VariableString & " Low Pass Filter"
            
            Comments = "Press Blue Button to turn " & VariableString & " Low Pass Filter"
            
            AdditComments = ""

            Hookup "ButtonAction", "", "", "", "", "", "", "", ""
'--------------------------------------End Button Press example Low Pass Filter Off-----------------------------
        
        
        ElseIf TestSect = 9000 Then
'--------------------------------------Begin Basic OpCheck-----------------------------
            
            ImageName = "1M"
            'ImageName = ImageNameString
            
            'SectTitle = "B79"
            SectTitle = SectTitleString
            
            'Title = "Insulation External Sense Test"
            Title = TitleString
            'Comments = "Turn knob to 1000V, Apply 35V/60Hz. Does display show >30V?"
            Comments = CommentsString
            AdditComments = ""

            Hookup "OpCheck", "", "", "", "", "", "", "", ""
'--------------------------------------End HRS Box Standard-----------------------------
     
        ElseIf TestSect = 9500 Then
'--------------------------------------Begin PassFail Is Reading Between 2 Values-----------------------------
            
            ImageName = "1M"
            'ImageName = ImageNameString
            
            'SectTitle = "B79"
            SectTitle = SectTitleString
            
            'Title = "Insulation External Sense Test"
            Title = TitleString
            'Comments = "Turn knob to 1000V, Apply 35V/60Hz. Does display show >30V?"
            Comments = CommentsString
            AdditComments = ""

            Hookup "OpCheck", "", "", "", "", "", "", "", ""
        
        ElseIf TestSect = 10000 Then
'--------------------------------------Begin PassFail Is Reading Between 2 Values-----------------------------
            
            ImageName = ""
            
            SectTitle = "B10000"
            
            Title = "Continuity Test"
            
            Comments = "Is the Continuity Beeper " & VariableString & "?"
            
            AdditComments = ""

            Hookup "OpCheck", "", "", "", "", "", "", "", ""
'--------------------------------------End Continuity Check-----------------------------
            
            
            
        End If

            
    
    
    
   ' If Dir(imagePath) <> "" Then
    '    Me.Image1.Picture = LoadPicture(imagePath)
   ' Else
    '    MsgBox "Image not found: " & imagePath, vbExclamation
   ' End If

End Sub

Private Sub Hookup(ByVal StateOn As String, UnitJack1 As String, STDJack1 As String, UnitJack2 As String, STDJack2 As String, UnitJack3 As String, STDJack3 As String, UnitJack4 As String, STDJack4 As String)
SetupWS
    
    Select Case StateOn
    
        Case "Calibrator"
            'MsgBox SectTitle
            'MsgBox DiffTitle
            If DiffTitle <> "" Then
                Me.Caption = Worksheets(Tab1).Range(DiffTitle).Value
            Else
                Me.Caption = Worksheets(Tab1).Range(SectTitle).Value
            End If
            Me.DMMComments.Visible = False
            Me.DMMAdditComments.Visible = False
            Me.DMMProbeLabel.Visible = False
            Me.Label34.Visible = True
            Me.Label31.Visible = True
            Me.Label28.Visible = True
            Me.Label25.Visible = True
            Me.Label23.Visible = True
            Me.Label33.Visible = True
            Me.HRSTextInput.Visible = False
            Me.HalfPageLabel.Visible = False
            Me.OpCheckMain.Visible = False
            Me.OpCheckComments.Visible = False
            Me.Label11.Caption = Model              'UUT Model
            Me.Label12.Caption = "<---->"
            Me.Label13.Caption = CalibratorModel    'Calibrator Model
            Me.Label32.Caption = UnitJack1          'Wiring from Meter
            Me.Label31.Caption = "<---->"           '<---->  Hookup between Unit and standards
            Me.Label30.Caption = STDJack1           'Wiring to Calibrator/DMM
            Me.Label29.Caption = UnitJack2          'Wiring from Meter
            Me.Label28.Caption = "<---->"           '<---->  Hookup between Unit and standards
            Me.Label27.Caption = STDJack2           'Wiring to Calibrator/DMM
            Me.Label26.Caption = UnitJack3                 'Wiring from Meter
            Me.Label25.Caption = "<---->"                 '<---->  Hookup between Unit and standards
            Me.Label24.Caption = STDJack3                 'Wiring to Calibrator/DMM
            Me.Label14.Caption = UnitJack4                 'Wiring from Meter
            Me.Label23.Caption = "<---->"                 '<---->  Hookup between Unit and standards
            Me.Label22.Caption = STDJack4                 'Wiring to Calibrator/DMM
            Me.Label6.Caption = Title
            Me.Label33.Caption = Comments
            Me.Label34.Caption = AdditComments
            Me.SourceMain.Visible = False
            Me.SourceMainComments.Visible = False
            Me.TextInput.Visible = False
            
            imagePath = ThisWorkbook.Path & "\Images\" & Model & "\" & CalibratorModel & "\" & ImageName & ".jpg"
    
        Case "UnitSource"
            If DiffTitle <> "" Then
                Me.Caption = Worksheets(Tab1).Range(DiffTitle).Value
            Else
                Me.Caption = Worksheets(Tab1).Range(SectTitle).Value
            End If
            Me.DMMComments.Visible = False
            Me.DMMAdditComments.Visible = False
            Me.Label34.Visible = True
            Me.Label31.Visible = True
            Me.Label28.Visible = True
            Me.Label25.Visible = True
            Me.Label23.Visible = True
            Me.Label33.Visible = True
            Me.HRSTextInput.Visible = False
            Me.HalfPageLabel.Visible = False
            Me.OpCheckMain.Visible = False
            Me.OpCheckComments.Visible = False
            Me.Label11.Caption = Model              'UUT Model
            Me.Label12.Caption = "<---->"
            Me.Label13.Caption = CalibratorModel    'Calibrator Model
            Me.Label32.Caption = UnitJack1          'Wiring from Meter
            Me.Label31.Caption = "<---->"           '<---->  Hookup between Unit and standards
            Me.Label30.Caption = STDJack1           'Wiring to Calibrator/DMM
            Me.Label29.Caption = UnitJack2          'Wiring from Meter
            Me.Label28.Caption = "<---->"           '<---->  Hookup between Unit and standards
            Me.Label27.Caption = STDJack2           'Wiring to Calibrator/DMM
            Me.Label26.Caption = UnitJack3                 'Wiring from Meter
            Me.Label25.Caption = "<---->"                 '<---->  Hookup between Unit and standards
            Me.Label24.Caption = STDJack3                 'Wiring to Calibrator/DMM
            Me.Label14.Caption = UnitJack4                 'Wiring from Meter
            Me.Label23.Caption = "<---->"                 '<---->  Hookup between Unit and standards
            Me.Label22.Caption = STDJack4                 'Wiring to Calibrator/DMM
            Me.Label6.Caption = Title
            Me.Label33.Caption = Comments
            Me.Label34.Caption = AdditComments
            Me.SourceMain.Visible = False
            Me.SourceMainComments.Visible = False
            Me.TextInput.Visible = False
            
            
            imagePath = ThisWorkbook.Path & "\Images\" & Model & "\" & CalibratorModel & "\" & ImageName & " " & CalibratorModel & ".jpg"

            
        Case "DMM"
            If DiffTitle <> "" Then
                Me.Caption = Worksheets(Tab1).Range(DiffTitle).Value
            Else
                Me.Caption = Worksheets(Tab1).Range(SectTitle).Value
            End If
            Me.DMMComments.Visible = True
            Me.DMMAdditComments.Visible = True
            Me.DMMProbeLabel.Visible = False
            Me.Label34.Visible = False
            Me.Label31.Visible = False
            Me.Label28.Visible = False
            Me.Label25.Visible = False
            Me.Label23.Visible = False
            Me.Label33.Visible = False
            Me.HRSTextInput.Visible = False
            Me.HalfPageLabel.Visible = False
            Me.OpCheckMain.Visible = False
            Me.OpCheckComments.Visible = False
            Me.Label11.Caption = Model              'UUT Model
            Me.Label12.Caption = "<---->"
            Me.Label13.Caption = DMMModel    'Calibrator Model
            Me.Label32.Caption = UnitJack1          'Wiring from Meter
            Me.Label31.Caption = "<---->"           '<---->  Hookup between Unit and standards
            Me.Label30.Caption = STDJack1           'Wiring to Calibrator/DMM
            Me.Label29.Caption = UnitJack2          'Wiring from Meter
            Me.Label28.Caption = "<---->"           '<---->  Hookup between Unit and standards
            Me.Label27.Caption = STDJack2           'Wiring to Calibrator/DMM
            Me.Label26.Caption = UnitJack3                 'Wiring from Meter
            Me.Label25.Caption = "<---->"                 '<---->  Hookup between Unit and standards
            Me.Label24.Caption = STDJack3                 'Wiring to Calibrator/DMM
            Me.Label14.Caption = UnitJack4                 'Wiring from Meter
            Me.Label23.Caption = "<---->"                 '<---->  Hookup between Unit and standards
            Me.Label22.Caption = STDJack4                 'Wiring to Calibrator/DMM
            Me.Label6.Caption = Title
            Me.DMMComments.Caption = Comments
            Me.DMMAdditComments.Caption = AdditComments
            Me.SourceMain.Visible = False
            Me.SourceMainComments.Visible = False
            Me.TextInput.Visible = False
            imagePath = ThisWorkbook.Path & "\Images\" & Model & "\" & DMMModel & "\" & ImageName & ".jpg"
            'DMMProbeLabel

            
        Case "DMMProbe"
            If DiffTitle <> "" Then
                Me.Caption = Worksheets(Tab1).Range(DiffTitle).Value
            Else
                Me.Caption = Worksheets(Tab1).Range(SectTitle).Value
            End If
            Me.DMMComments.Visible = False
            Me.DMMAdditComments.Visible = False
            Me.Label34.Visible = False
            Me.Label31.Visible = False
            Me.Label28.Visible = False
            Me.Label25.Visible = False
            Me.Label23.Visible = False
            Me.Label33.Visible = False
            Me.HRSTextInput.Visible = False
            Me.HalfPageLabel.Visible = False
            Me.OpCheckMain.Visible = False
            Me.OpCheckComments.Visible = False
            Me.DMMProbeLabel.Visible = True
            Me.Label11.Caption = Model              'UUT Model
            Me.Label12.Caption = "<---->"
            Me.Label13.Caption = DMMModel    'Calibrator Model
            Me.Label32.Caption = UnitJack1          'Wiring from Meter
            Me.Label31.Caption = "<---->"           '<---->  Hookup between Unit and standards
            Me.Label30.Caption = STDJack1           'Wiring to Calibrator/DMM
            Me.Label29.Caption = UnitJack2          'Wiring from Meter
            Me.Label28.Caption = "<---->"           '<---->  Hookup between Unit and standards
            Me.Label27.Caption = STDJack2           'Wiring to Calibrator/DMM
            Me.Label26.Caption = UnitJack3                 'Wiring from Meter
            Me.Label25.Caption = "<---->"                 '<---->  Hookup between Unit and standards
            Me.Label24.Caption = STDJack3                 'Wiring to Calibrator/DMM
            Me.Label14.Caption = UnitJack4                 'Wiring from Meter
            Me.Label23.Caption = "<---->"                 '<---->  Hookup between Unit and standards
            Me.Label22.Caption = STDJack4                 'Wiring to Calibrator/DMM
            Me.Label6.Caption = Title
            Me.DMMProbeLabel.Caption = Comments
            Me.SourceMain.Visible = False
            Me.SourceMainComments.Visible = False
            Me.TextInput.Visible = False
            imagePath = ThisWorkbook.Path & "\Images\" & Model & "\" & DMMModel & "\" & ImageName & " " & DMMModel & ".jpg"
            'DMMProbeLabel
            
        Case "OpCheck"
        
            If DiffTitle <> "" Then
                Me.Caption = Worksheets(Tab1).Range(DiffTitle).Value
            Else
                Me.Caption = Worksheets(Tab1).Range(SectTitle).Value
            End If
            Me.DMMComments.Visible = False
            Me.DMMAdditComments.Visible = False
            Me.Label34.Visible = False
            Me.Label31.Visible = False
            Me.Label28.Visible = False
            Me.Label25.Visible = False
            Me.Label23.Visible = False
            Me.Label33.Visible = False
            Me.Label11.Visible = False              'UUT Model
            Me.Label12.Visible = False
            Me.Label13.Visible = False    'Calibrator Model
            Me.Label32.Visible = False          'Wiring from Meter
            Me.Label31.Visible = False           '<---->  Hookup between Unit and standards
            Me.Label30.Visible = False           'Wiring to Calibrator/DMM
            Me.Label29.Visible = False          'Wiring from Meter
            Me.Label28.Visible = False           '<---->  Hookup between Unit and standards
            Me.Label27.Visible = False           'Wiring to Calibrator/DMM
            Me.Label26.Visible = False                 'Wiring from Meter
            Me.Label25.Visible = False                 '<---->  Hookup between Unit and standards
            Me.Label24.Visible = False                 'Wiring to Calibrator/DMM
            Me.Label14.Visible = False                 'Wiring from Meter
            Me.Label23.Visible = False                 '<---->  Hookup between Unit and standards
            Me.Label22.Visible = False                 'Wiring to Calibrator/DMM
            Me.Label6.Visible = False
            Me.DMMProbeLabel.Visible = False
            Me.HRSTextInput.Visible = False
            Me.HalfPageLabel.Visible = False
            Me.OpCheckMain.Visible = True
            Me.OpCheckMain.Caption = Comments
            Me.OpCheckComments.Visible = True
            Me.OpCheckComments.Caption = AdditComments
            Me.Pass.Visible = True
            Me.Fail.Visible = True
            Me.Advance.Visible = False
            Me.SourceMain.Visible = False
            Me.SourceMainComments.Visible = False
            Me.TextInput.Visible = False
            'imagePath = ThisWorkbook.Path & "\Images\" & Model & "\" & DMMModel & "\" & ImageName & " " & DMMModel & ".jpg"

            
        Case "Stabilize"
            If DiffTitle <> "" Then
                Me.Caption = Worksheets(Tab1).Range(DiffTitle).Value
            Else
                Me.Caption = Worksheets(Tab1).Range(SectTitle).Value
            End If
            Me.DMMComments.Visible = False
            Me.DMMAdditComments.Visible = False
            Me.Label34.Visible = False
            Me.Label31.Visible = False
            Me.Label28.Visible = False
            Me.Label25.Visible = False
            Me.Label23.Visible = False
            Me.Label33.Visible = False
            Me.Label11.Visible = False              'UUT Model
            Me.Label12.Visible = False
            Me.Label13.Visible = False    'Calibrator Model
            Me.Label32.Visible = False          'Wiring from Meter
            Me.Label31.Visible = False           '<---->  Hookup between Unit and standards
            Me.Label30.Visible = False           'Wiring to Calibrator/DMM
            Me.Label29.Visible = False          'Wiring from Meter
            Me.Label28.Visible = False           '<---->  Hookup between Unit and standards
            Me.Label27.Visible = False           'Wiring to Calibrator/DMM
            Me.Label26.Visible = False                 'Wiring from Meter
            Me.Label25.Visible = False                 '<---->  Hookup between Unit and standards
            Me.Label24.Visible = False                 'Wiring to Calibrator/DMM
            Me.Label14.Visible = False                 'Wiring from Meter
            Me.Label23.Visible = False                 '<---->  Hookup between Unit and standards
            Me.Label22.Visible = False                 'Wiring to Calibrator/DMM
            Me.Label6.Visible = False
            Me.HalfPageLabel.Visible = False
            Me.HRSTextInput.Visible = False
            Me.DMMProbeLabel.Visible = False
            Me.OpCheckMain.Visible = True
            Me.OpCheckMain.Caption = Comments
            Me.OpCheckComments.Visible = True
            Me.OpCheckComments.Caption = AdditComments
            Me.Pass.Visible = False
            Me.Fail.Visible = False
            Me.Advance.Visible = True
            Me.SourceMain.Visible = False
            Me.SourceMainComments.Visible = False
            Me.TextInput.Visible = False
            'imagePath = ThisWorkbook.Path & "\Images\" & Model & "\" & DMMModel & "\" & ImageName & " " & DMMModel & ".jpg"

            
        Case "Sourcing"
            
            If DiffTitle <> "" Then
                Me.Caption = Worksheets(Tab1).Range(DiffTitle).Value
            Else
                Me.Caption = Worksheets(Tab1).Range(SectTitle).Value
            End If
            Me.DMMComments.Visible = False
            Me.DMMAdditComments.Visible = False
            Me.Label34.Visible = False
            Me.Label31.Visible = False
            Me.Label28.Visible = False
            Me.Label25.Visible = False
            Me.Label23.Visible = False
            Me.Label33.Visible = False
            Me.Label11.Visible = False              'UUT Model
            Me.Label12.Visible = False
            Me.Label13.Visible = False    'Calibrator Model
            Me.Label32.Visible = False          'Wiring from Meter
            Me.Label31.Visible = False           '<---->  Hookup between Unit and standards
            Me.Label30.Visible = False           'Wiring to Calibrator/DMM
            Me.Label29.Visible = False          'Wiring from Meter
            Me.Label28.Visible = False           '<---->  Hookup between Unit and standards
            Me.Label27.Visible = False           'Wiring to Calibrator/DMM
            Me.Label26.Visible = False                 'Wiring from Meter
            Me.Label25.Visible = False                 '<---->  Hookup between Unit and standards
            Me.Label24.Visible = False                 'Wiring to Calibrator/DMM
            Me.Label14.Visible = False                 'Wiring from Meter
            Me.Label23.Visible = False                 '<---->  Hookup between Unit and standards
            Me.Label22.Visible = False                 'Wiring to Calibrator/DMM
            Me.Label6.Visible = False
            Me.DMMProbeLabel.Visible = False
            Me.HRSTextInput.Visible = False
            Me.HalfPageLabel.Visible = False
            Me.OpCheckMain.Visible = False
            Me.SourceMain.Visible = True
            Me.SourceMain.Caption = Comments
            Me.SourceMainComments.Visible = True
            Me.SourceMainComments.Caption = AdditComments
            Me.OpCheckComments.Visible = False
            Me.Pass.Visible = False
            Me.Fail.Visible = False
            Me.Advance.Visible = True
            Me.TextInput.Visible = False
            'imagePath = ThisWorkbook.Path & "\Images\" & Model & "\" & DMMModel & "\" & ImageName & " " & DMMModel & ".jpg"

            
        Case "SourcingTextInput"
            
            If DiffTitle <> "" Then
                Me.Caption = Worksheets(Tab1).Range(DiffTitle).Value
            Else
                Me.Caption = Worksheets(Tab1).Range(SectTitle).Value
            End If
            Me.DMMComments.Visible = False
            Me.DMMAdditComments.Visible = False
            Me.Label34.Visible = False
            Me.Label31.Visible = False
            Me.Label28.Visible = False
            Me.Label25.Visible = False
            Me.Label23.Visible = False
            Me.Label33.Visible = False
            Me.Label11.Visible = False              'UUT Model
            Me.Label12.Visible = False
            Me.Label13.Visible = False    'Calibrator Model
            Me.Label32.Visible = False          'Wiring from Meter
            Me.Label31.Visible = False           '<---->  Hookup between Unit and standards
            Me.Label30.Visible = False           'Wiring to Calibrator/DMM
            Me.Label29.Visible = False          'Wiring from Meter
            Me.Label28.Visible = False           '<---->  Hookup between Unit and standards
            Me.Label27.Visible = False           'Wiring to Calibrator/DMM
            Me.Label26.Visible = False                 'Wiring from Meter
            Me.Label25.Visible = False                 '<---->  Hookup between Unit and standards
            Me.Label24.Visible = False                 'Wiring to Calibrator/DMM
            Me.Label14.Visible = False                 'Wiring from Meter
            Me.Label23.Visible = False                 '<---->  Hookup between Unit and standards
            Me.Label22.Visible = False                 'Wiring to Calibrator/DMM
            Me.Label6.Visible = False
            Me.DMMProbeLabel.Visible = False
            Me.HRSTextInput.Visible = False
            Me.HalfPageLabel.Visible = False
            Me.OpCheckMain.Visible = False
            Me.SourceMain.Visible = True
            Me.SourceMain.Caption = Comments
            Me.SourceMainComments.Visible = True
            Me.SourceMainComments.Caption = AdditComments
            Me.OpCheckComments.Visible = False
            Me.Pass.Visible = False
            Me.Fail.Visible = False
            Me.Advance.Visible = True
            Me.TextInput.Visible = True
            'imagePath = ThisWorkbook.Path & "\Images\" & Model & "\" & DMMModel & "\" & ImageName & " " & DMMModel & ".jpg"

            
        Case "ButtonAction"
            
            If DiffTitle <> "" Then
                Me.Caption = Worksheets(Tab1).Range(DiffTitle).Value
            Else
                Me.Caption = Worksheets(Tab1).Range(SectTitle).Value
            End If
            Me.DMMComments.Visible = False
            Me.DMMAdditComments.Visible = False
            Me.Label34.Visible = False
            Me.Label31.Visible = False
            Me.Label28.Visible = False
            Me.Label25.Visible = False
            Me.Label23.Visible = False
            Me.Label33.Visible = False
            Me.Label11.Visible = False              'UUT Model
            Me.Label12.Visible = False
            Me.Label13.Visible = False    'Calibrator Model
            Me.Label32.Visible = False          'Wiring from Meter
            Me.Label31.Visible = False           '<---->  Hookup between Unit and standards
            Me.Label30.Visible = False           'Wiring to Calibrator/DMM
            Me.Label29.Visible = False          'Wiring from Meter
            Me.Label28.Visible = False           '<---->  Hookup between Unit and standards
            Me.Label27.Visible = False           'Wiring to Calibrator/DMM
            Me.Label26.Visible = False                 'Wiring from Meter
            Me.Label25.Visible = False                 '<---->  Hookup between Unit and standards
            Me.Label24.Visible = False                 'Wiring to Calibrator/DMM
            Me.Label14.Visible = False                 'Wiring from Meter
            Me.Label23.Visible = False                 '<---->  Hookup between Unit and standards
            Me.Label22.Visible = False                 'Wiring to Calibrator/DMM
            Me.Label6.Visible = False
            Me.DMMProbeLabel.Visible = False
            Me.HRSTextInput.Visible = False
            Me.HalfPageLabel.Visible = False
            Me.OpCheckMain.Visible = False
            Me.SourceMain.Visible = True
            Me.SourceMain.Caption = Comments
            Me.SourceMainComments.Visible = True
            Me.SourceMainComments.Caption = AdditComments
            Me.OpCheckComments.Visible = False
            Me.Pass.Visible = False
            Me.Fail.Visible = False
            Me.Advance.Visible = True
            Me.TextInput.Visible = False
            'imagePath = ThisWorkbook.Path & "\Images\" & Model & "\" & DMMModel & "\" & ImageName & " " & DMMModel & ".jpg"
    
        Case "HRS"
            If DiffTitle <> "" Then
                Me.Caption = Worksheets(Tab1).Range(DiffTitle).Value
            Else
                Me.Caption = Worksheets(Tab1).Range(SectTitle).Value
            End If
            Me.DMMComments.Visible = False
            Me.DMMProbeLabel.Visible = False
            Me.DMMAdditComments.Visible = False
            Me.Label34.Visible = False
            Me.Label31.Visible = True
            Me.Label28.Visible = True
            Me.Label25.Visible = True
            Me.Label23.Visible = True
            Me.Label33.Visible = False
            Me.HalfPageLabel.Visible = True
            Me.OpCheckMain.Visible = False
            Me.OpCheckComments.Visible = False
            Me.Label11.Caption = Model              'UUT Model
            Me.Label12.Caption = "<---->"
            Me.Label13.Caption = CalibratorModel    'Calibrator Model
            Me.Label32.Caption = UnitJack1          'Wiring from Meter
            Me.Label31.Caption = "<---->"           '<---->  Hookup between Unit and standards
            Me.Label30.Caption = STDJack1           'Wiring to Calibrator/DMM
            Me.Label29.Caption = UnitJack2          'Wiring from Meter
            Me.Label28.Caption = "<---->"           '<---->  Hookup between Unit and standards
            Me.Label27.Caption = STDJack2           'Wiring to Calibrator/DMM
            Me.Label26.Caption = UnitJack3                 'Wiring from Meter
            Me.Label25.Caption = "<---->"                 '<---->  Hookup between Unit and standards
            Me.Label24.Caption = STDJack3                 'Wiring to Calibrator/DMM
            Me.Label14.Caption = UnitJack4                 'Wiring from Meter
            Me.Label23.Caption = "<---->"                 '<---->  Hookup between Unit and standards
            Me.Label22.Caption = STDJack4                 'Wiring to Calibrator/DMM
            Me.Label6.Caption = Title
            Me.Label33.Caption = Comments
            Me.Label34.Caption = AdditComments
            Me.SourceMain.Visible = False
            Me.SourceMainComments.Visible = False
            Me.TextInput.Visible = False
            Me.HRSTextInput.Visible = True
            Me.HRSTextInput.SetFocus
            Me.HalfPageLabel.Caption = Comments
            
            imagePath = ThisWorkbook.Path & "\Images\" & Model & "\HRS\" & ImageName & " " & "HRS.jpg"
            
    End Select
    If Dir(imagePath) <> "" Then
        Me.Image1.Picture = LoadPicture(imagePath)
        
    Else
        MsgBox "Image not found: " & imagePath, vbExclamation
    End If
   
End Sub

Private Sub HRSTextInput_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then   ' Enter key pressed
        KeyCode = 0                 ' Prevents the "ding" sound
        Call Advance_Click   ' Run your Advance button code
    End If
End Sub




