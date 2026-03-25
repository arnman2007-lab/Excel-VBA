Attribute VB_Name = "SetupArrays"
Sub ArraySetup()

 ' Examples
    'If there is a space between same type of test points i.e. AC Voltage, then each section of test points will be a new TestSect
    'Or if there is a change, i.e. Range change or press a button, then make a new TestSect
    'TestPointVoltage = Array(5, 50, 50, 250, 500)
    'TestPointFrequency = Array(20, 65, 100, 65, 45)
    'TestPointVoltageUnits = Array(V, V, mV, mV, V)
    'TestPointFrequencyUnits = Array("Hz", "kHz", "kHz", "kHz", "Hz")
    
    ranges = Array("20:26", "29:37", "40:51", "54:54", "55:55", "58:58", "59:59", "61:64")
    Skips = Array("27:27", "38:38", "52:52", "56:56")
    stdbyComms = Array("28:28", "39:39", "53:53", "57:57")
    
    'Input The Last gray Cell address, right after the last Reading input cell, i.e. $F$89 for data input columns
    'For Example you have F and G as your as Found and As Left columns and the last row/cell where there are no more
    'cell to input data and that next cell is gray color, or whatever color it is. Be sure to include the $'s.
    
    'LastCellF = "$F$89"
    'LastCellG = "$G$89"
    'LastCellH = ""
    
    LastCellF = ""
    LastCellG = "$G$65"
    LastCellH = "$H$65"
'Section 1
If TestSect = 1 Then
TestPoint = Array(0.1, 0.1, 0.1, 0.1, 0.1, 0.1, 0.1)
TestPointUnits = Array("V", "V", "V", "V", "V", "V", "V")
TestPointFrequency = Array(10, 50, 100, 500, 1, 5, 10)
TestPointFrequencyUnits = Array("Hz", "Hz", "Hz", "Hz", "kHz", "kHz", "kHz")
TestPointWave = Array()
TestPointOffset = Array()
TestPointComp = Array()
TestPointDuty = Array()

'SameTest uses the same hookups and meter configuration i.e. changing ranges. SameTest is used to minimize userform popups between same tests that are seperated by the skips.
'If two sections of testpoints are seperated by a gray blank, and have the same setup, then use the same SameTest number. If the next test changes, where the user needs to change something,
'then give SameTest a different number for every different setup.

SameTest = 1

'Section 2
ElseIf TestSect = 2 Then
TestPoint = Array(0.005, 0.01, 0.05, 0.1, 0.2, 0.5, 1, 1.5, 1.95)
TestPointUnits = Array("V", "V", "V", "V", "V", "V", "V", "V", "V")
TestPointFrequency = Array(100, 100, 100, 100, 100, 100, 100, 100, 100)
TestPointFrequencyUnits = Array("Hz", "Hz", "Hz", "Hz", "Hz", "Hz", "Hz", "Hz", "Hz")
TestPointWave = Array()
TestPointOffset = Array()
TestPointComp = Array()
TestPointDuty = Array()

SameTest = 1

ElseIf TestSect = 3 Then
'Section 3
TestPoint = Array(0.1, 0.1, 0.1, 0.2, 0.2, 0.2, 0.3, 0.6, 1.2, 1.5, 1.9, 1.9)
TestPointUnits = Array("V", "V", "V", "V", "V", "V", "V", "V", "V", "V", "V", "V")
TestPointFrequency = Array(10, 15, 20, 25, 30, 50, 100, 500, 1, 2, 5, 10)
TestPointFrequencyUnits = Array("Hz", "Hz", "Hz", "Hz", "Hz", "Hz", "Hz", "Hz", "kHz", "kHz", "kHz", "kHz")
TestPointWave = Array()
TestPointOffset = Array()
TestPointComp = Array()
TestPointDuty = Array()

SameTest = 2

ElseIf TestSect = 4 Then
'Section 4
TestPoint = Array(0.1)
TestPointUnits = Array("V")
TestPointFrequency = Array(500)
TestPointFrequencyUnits = Array("Hz")
TestPointWave = Array()
TestPointOffset = Array()
TestPointComp = Array()
TestPointDuty = Array()

SameTest = 3

ElseIf TestSect = 5 Then
'Section 5
TestPoint = Array(0.1)
TestPointUnits = Array("V")
TestPointFrequency = Array(5)
TestPointFrequencyUnits = Array("kHz")
TestPointWave = Array()
TestPointOffset = Array()
TestPointComp = Array()
TestPointDuty = Array()

SameTest = 3

ElseIf TestSect = 6 Then
'Section 6
TestPoint = Array(0.1)
TestPointUnits = Array("V")
TestPointFrequency = Array(440)
TestPointFrequencyUnits = Array("Hz")
TestPointWave = Array()
TestPointOffset = Array()
TestPointComp = Array()
TestPointDuty = Array()

SameTest = 4

ElseIf TestSect = 7 Then
'Section 7
TestPoint = Array()
TestPointUnits = Array()
TestPointFrequency = Array()
TestPointFrequencyUnits = Array()
TestPointWave = Array()
TestPointOffset = Array()
TestPointComp = Array()
TestPointDuty = Array()

SameTest = 6

ElseIf TestSect = 8 Then
'Section 8
TestPoint = Array()
TestPointUnits = Array()
TestPointFrequency = Array()
TestPointFrequencyUnits = Array()
TestPointWave = Array()
TestPointOffset = Array()
TestPointComp = Array()
TestPointDuty = Array()

SameTest = 7

ElseIf TestSect = 9 Then
'Section 9
TestPoint = Array()
TestPointUnits = Array()
TestPointFrequency = Array()
TestPointFrequencyUnits = Array()
TestPointWave = Array()
TestPointOffset = Array()
TestPointComp = Array()
TestPointDuty = Array()

SameTest = 7

ElseIf TestSect = 10 Then
'Section 10
TestPoint = Array()
TestPointUnits = Array()
TestPointFrequency = Array()
TestPointFrequencyUnits = Array()
TestPointWave = Array()
TestPointOffset = Array()
TestPointComp = Array()
TestPointDuty = Array()

SameTest = 7

ElseIf TestSect = 11 Then
'Section 11
TestPoint = Array()
TestPointUnits = Array()
TestPointFrequency = Array()
TestPointFrequencyUnits = Array()
TestPointWave = Array()
TestPointOffset = Array()
TestPointComp = Array()
TestPointDuty = Array()

SameTest = 7

ElseIf TestSect = 12 Then
'Section 12
TestPoint = Array()
TestPointUnits = Array()
TestPointFrequency = Array()
TestPointFrequencyUnits = Array()
TestPointWave = Array()
TestPointOffset = Array()
TestPointComp = Array()
TestPointDuty = Array()

SameTest = 8

ElseIf TestSect = 13 Then
'Section 13
TestPoint = Array()
TestPointUnits = Array()
TestPointFrequency = Array()
TestPointFrequencyUnits = Array()
TestPointWave = Array()
TestPointOffset = Array()
TestPointComp = Array()
TestPointDuty = Array()

SameTest = 8

ElseIf TestSect = 14 Then
'Section 14
TestPoint = Array()
TestPointUnits = Array()
TestPointFrequency = Array()
TestPointFrequencyUnits = Array()
TestPointWave = Array()
TestPointOffset = Array()
TestPointComp = Array()
TestPointDuty = Array()

SameTest = 10

ElseIf TestSect = 15 Then
'Section 15
TestPoint = Array()
TestPointUnits = Array()
TestPointFrequency = Array()
TestPointFrequencyUnits = Array()
TestPointWave = Array()
TestPointOffset = Array()
TestPointComp = Array()
TestPointDuty = Array()

SameTest = 11


ElseIf TestSect = 16 Then
'Section 16
TestPoint = Array()
TestPointUnits = Array()
TestPointFrequency = Array()
TestPointFrequencyUnits = Array()
TestPointWave = Array()
TestPointOffset = Array()
TestPointComp = Array()
TestPointDuty = Array()

SameTest = 12

ElseIf TestSect = 17 Then
'Section 17
TestPoint = Array()
TestPointUnits = Array()
TestPointFrequency = Array()
TestPointFrequencyUnits = Array()
TestPointWave = Array()
TestPointOffset = Array()
TestPointComp = Array()
TestPointDuty = Array()

SameTest = 13

ElseIf TestSect = 18 Then
'Section 18
TestPoint = Array()
TestPointUnits = Array()
TestPointFrequency = Array()
TestPointFrequencyUnits = Array()
TestPointWave = Array()
TestPointOffset = Array()
TestPointComp = Array()
TestPointDuty = Array()

SameTest = 14

ElseIf TestSect = 19 Then
'Section 19
TestPoint = Array()
TestPointUnits = Array()
TestPointFrequency = Array()
TestPointFrequencyUnits = Array()
TestPointWave = Array()
TestPointOffset = Array()
TestPointComp = Array()
TestPointDuty = Array()

SameTest = 15

ElseIf TestSect = 20 Then
'Section 20
TestPoint = Array()
TestPointUnits = Array()
TestPointFrequency = Array()
TestPointFrequencyUnits = Array()
TestPointWave = Array()
TestPointOffset = Array()
TestPointComp = Array()
TestPointDuty = Array()

ElseIf TestSect = 21 Then
'Section 21
TestPoint = Array()
TestPointUnits = Array()
TestPointFrequency = Array()
TestPointFrequencyUnits = Array()
TestPointWave = Array()
TestPointOffset = Array()
TestPointComp = Array()
TestPointDuty = Array()

ElseIf TestSect = 22 Then
'Section 22
TestPoint = Array()
TestPointUnits = Array()
TestPointFrequency = Array()
TestPointFrequencyUnits = Array()
TestPointWave = Array()
TestPointOffset = Array()
TestPointComp = Array()
TestPointDuty = Array()

ElseIf TestSect = 23 Then
'Section 23
TestPoint = Array()
TestPointUnits = Array()
TestPointFrequency = Array()
TestPointFrequencyUnits = Array()
TestPointWave = Array()
TestPointOffset = Array()
TestPointComp = Array()
TestPointDuty = Array()

ElseIf TestSect = 24 Then
'Section 24
TestPoint = Array()
TestPointUnits = Array()
TestPointFrequency = Array()
TestPointFrequencyUnits = Array()
TestPointWave = Array()
TestPointOffset = Array()
TestPointComp = Array()
TestPointDuty = Array()

ElseIf TestSect = 25 Then
'Section 25
TestPoint = Array()
TestPointUnits = Array()
TestPointFrequency = Array()
TestPointFrequencyUnits = Array()
TestPointWave = Array()
TestPointOffset = Array()
TestPointComp = Array()
TestPointDuty = Array()

ElseIf TestSect = 26 Then
'Section 26
TestPoint = Array()
TestPointUnits = Array()
TestPointFrequency = Array()
TestPointFrequencyUnits = Array()
TestPointWave = Array()
TestPointOffset = Array()
TestPointComp = Array()
TestPointDuty = Array()

ElseIf TestSect = 27 Then
'Section 27
TestPoint = Array()
TestPointUnits = Array()
TestPointFrequency = Array()
TestPointFrequencyUnits = Array()
TestPointWave = Array()
TestPointOffset = Array()
TestPointComp = Array()
TestPointDuty = Array()

ElseIf TestSect = 28 Then
'Section 28
TestPoint = Array()
TestPointUnits = Array()
TestPointFrequency = Array()
TestPointFrequencyUnits = Array()
TestPointWave = Array()
TestPointOffset = Array()
TestPointComp = Array()
TestPointDuty = Array()

ElseIf TestSect = 29 Then
'Section 29
TestPoint = Array()
TestPointUnits = Array()
TestPointFrequency = Array()
TestPointFrequencyUnits = Array()
TestPointWave = Array()
TestPointOffset = Array()
TestPointComp = Array()
TestPointDuty = Array()

ElseIf TestSect = 30 Then
'Section 30
TestPoint = Array()
TestPointUnits = Array()
TestPointFrequency = Array()
TestPointFrequencyUnits = Array()
TestPointWave = Array()
TestPointOffset = Array()
TestPointComp = Array()
TestPointDuty = Array()
End If
End Sub
