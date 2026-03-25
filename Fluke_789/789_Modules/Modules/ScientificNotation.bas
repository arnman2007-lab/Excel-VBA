Attribute VB_Name = "ScientificNotation"
Sub FixReading(DMMQueryF As String)
        Dim scientificNumber As String
        Dim cleanedNumber As String
        'MsgBox "DMMQueryF Reading " & DMMQueryF
        
        scientificNumber = DMMQueryF
       ' MsgBox "scientificNumber Reading " & scientificNumber
       
        cleanedNumber = ExtractScientificNumber(scientificNumber)
        'MsgBox "cleanedNumber Reading " & cleanedNumber
        
        If Multiplier <> 0 Then
            FixedRdg = CDbl(Trim(cleanedNumber)) * Multiplier
            'MsgBox "Multiplied Fixed Reading " & FixedRdg
            
         ElseIf Divider <> 0 Then
            FixedRdg = CDbl(Trim(cleanedNumber)) / Divider
            'MsgBox "Divided Fixed Reading " & FixedRdg
            
         End If

End Sub

Function ExtractScientificNumber(inputString As String) As String
    Dim regex       As Object
    Set regex = CreateObject("VBScript.RegExp")
    
    With regex
        .Pattern = "[-+]?[0-9]*\.?[0-9]+([eE][-+]?[0-9]+)?"
        .Global = True
    End With
    
    If regex.Test(inputString) Then
        ExtractScientificNumber = regex.Execute(inputString)(0)
    Else
        ExtractScientificNumber = "0"
    End If
End Function

