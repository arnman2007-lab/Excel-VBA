Attribute VB_Name = "TestSectionNumbers"
Sub TestSections(TestSection As Double)

ButtonState PanelForm, "CodeButton", "Operating"
Select Case TestSection
    Case "1000"
'--------Temp Measure Stabilize----------
        TestSect = TestSection
    Case "2000"
'--------Frequency Source From Unit------
        TestSect = TestSection
    Case "3000"
'--------DC mV Source From Unit----------------------------
        TestSect = TestSection
    Case "4000"
'--------Begin DC V Source From Unit----------------------------
        TestSect = TestSection
    Case "5000"
'--------Begin Ohms Source From Unit----------------------------
        TestSect = TestSection
    Case "6000"
'--------Begin mA Source From Unit----------------------------
        TestSect = TestSection
    Case "7000"
'--------Begin Insulation Tests HRS Box-----------------------------
        TestSect = TestSection
    Case "8000"
'--------Button Press example Low Pass On/Off----------
        TestSect = TestSection
    Case "9000"
'--------Next Test Here----------
        TestSect = TestSection
    Case "9000"

        TestSect = TestSection
    Case "10000"
'--------Continuity Check----------
        TestSect = TestSection
    Case Else
        TestSect = TestSection
    
        
        TestSectBak = TestSection
End Select

End Sub
