Attribute VB_Name = "TestSectionNumbers"
' Fluke 771 Milliamp Process Clamp Meter - Test Section Numbers
' This meter only measures DC mA, so test sections are simplified

Sub TestSections(TestSection As Double)

ButtonState PanelForm, "CodeButton", "Operating"
Select Case TestSection
    Case "1000"
'--------Operational Checks (Backlight, Display, Keypad, Spotlight)----------
        TestSect = TestSection
    Case "6000"
'--------DC mA Source From Calibrator----------------------------
'        Calibrator sources mA, UUT clamps around lead to measure
        TestSect = TestSection
    Case Else
        TestSect = TestSection
        TestSectBak = TestSection
End Select

End Sub
