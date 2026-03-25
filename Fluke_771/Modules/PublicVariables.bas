Attribute VB_Name = "PublicVariables"
Option Explicit
    
Public usbcommvalue As Range
Public ioMgr       As VisaComLib.ResourceManager
'Public ioMgr       As Object
Public CalibDevice  As VisaComLib.FormattedIO488
Public instru      As VisaComLib.FormattedIO488
Public addstdvalue As Range
Public gpib        As VisaComLib.IGpib
Public gpib3458    As VisaComLib.IGpib
Public DMMDevice    As VisaComLib.FormattedIO488
Public CounterDevice    As VisaComLib.FormattedIO488
Public add3458value As Range
Public addCommvalue As Range
Public Initialized3458 As String
Public InitializedCalibrator As String
Public ioMgrSet As Boolean
Public shouldContinue As Boolean
Public TerminateClicked As Boolean
Public PassClicked As Boolean
Public FailClicked As Boolean
Public addresses As Variant
Public func         As String
Public funct        As String
Public inst_valueFinal As Double
Public Tab1 As String
Public Tab2 As String
Public Tab3 As String
Public Tab4 As String
Public ATab1 As String
Public ATab2 As String
Public ATab3 As String
Public ATab4 As String
Public WorkOrder As String
Public Accredited As String
Public PathtoWB As String
Public dataSheet As Worksheet
Public InfoDataSheet As Worksheet
Public InfoSheet As Worksheet
Public WorkOrderSheet As Worksheet
Public cellAddress As Range
Public ranges As Object
Public ranges2 As Variant
Public ranges3 As Variant
Public PrevAddress As String
Public PrevTestSect As Integer
Public SameTest As Integer
Public PrevSameTest As Integer
Public OffValueV As Double
Public OffValueU As String
Public OffValueHz As Double
Public OffValueHzU As String
Public OffSet As Double
Public OffSetU As String
Public Wave As String
Public CalibListed As Boolean
Public FreqUser As Boolean
Public TestSect As Integer
Public Rangelck As String
Public CalibStatus As String
Public make As String
Public Model As String
Public UnitDesc As String
Public Check As String
Public comp As String
Public Duty As Double
Public activeCol As Long
Public inputString As String
Public i As Integer
Public found As Boolean
Public prefix As String
Public ohmString As String
Public numericValue As String
Public tabName As String
Public Skips As Variant
Public stdbyComms As Variant
Public TestSectA As Variant
Public HVImageLoc As String
Public TestPoint As Variant
Public TestPointUnits As Variant
Public TestPointFrequency As Variant
Public TestPointFrequencyUnits As Variant
Public TestPointWave As Variant
Public TestPointOffset As Variant
Public TestPointComp As Variant
Public TestPointDuty As Variant
Public LastCellF As String
Public LastCellG As String
Public LastCellH As String
Public ColNumAF As Integer
Public ColNumAL As Integer
Public TestForm As Integer
Public RdgRange As Double
Public gpibDevices() As String
Public asrlDevices() As String

'''''''''''''''DMM Stuff Begin'''''''''''''''''
Public NPLCNumber As Double
Public NRDGSNumber As Double
Public TrigEvent As String
'''''''''''''''DMM Stuff End""'''''''''''''''''

'''''''''''''''Values from 1WorkStation Setup.xlsm Begin'''''''''''''''
Public CalibratorMake As String
Public CalibratorModel As String
Public CalibratorSN As String
Public CalibratorGPIB As String
Public CalibratorScopeOption As String
Public DMMMake As String
Public DMMModel As String
Public DMMSN As String
Public DMMGPIB As String
Public CounterMake As String
Public CounterModel As String
Public CounterSN As String
Public CounterGPIB As String
'''''''''''''''Values from 1WorkStation Setup.xlsm End  '''''''''''''''

'''''''''''''''Begin Calibrator Arguments '''''''''''''''''''''''''''''
Public CalibMode As String
Public CalibCalFunc As String
Public CalibParam As Double
Public CalibParamUnit As String
Public CalibHertz As Double
Public CalibHertzUnit As String
Public CalibWave As String
Public CalibOffSet As Double
Public CalibDuty As Double
Public CalibZComp As String
'''''''''''''''End Calibrator Arguments '''''''''''''''''''''''''''''
Public CanDoIt As Integer
Public ready As Integer

'-------------------Begin Standard Reset Check Variable Setup----------------------------------
Public CalibratorReset As Integer
Public DMMReset As Integer
Public CounterReset As Integer
'-------------------End Standard Reset Check Variable Setup------------------------------------

Public DMMQuery As String

'-------------------Begin Button Functions-----------------------------------------------------
Public PreloadCols() As String
'-------------------End Button Functions-----------------------------------------------------
Public OpDone As Double
Public TestSectBak  As Double
Public NewSTD As String
Public NewHRS As String
Public VariableString As String
Public VariableDouble As Double
Public Multiplier As Double
Public Divider As Double
Public DiffTitle As String
Public FixedRdg As Double
Public AutoSelect As Boolean
Public ImageNameString As String
Public SectTitleString As String
Public TitleString As String
Public CommentsString As String
Public PassFailLimit1 As Double
Public PassFailLImit2 As Double


