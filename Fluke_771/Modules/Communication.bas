Attribute VB_Name = "Communication"


Sub ListCommPorts()
    On Error Resume Next

    Dim resources() As String
    Dim RM_P As VisaComLib.ResourceManager
    Dim resourceName As Variant
    Dim asrlResources As String
    Dim InfoSheet As Worksheet
    Dim commCell As Range

    ' Set reference to the Information sheet
    Set InfoSheet = ThisWorkbook.Sheets("Information")
    ' P15:Q15 merged cell
    Set commCell = InfoSheet.Range("P15")

    ' Clear existing content
    commCell.ClearContents

    ' Create VISA Resource Manager
    Set RM_P = New VisaComLib.ResourceManager

    ' Find all resources
    resources = RM_P.FindRsrc("?*")

    ' Always start with a friendly label
    asrlResources = "Click dropdown..."

    ' Loop through and filter ASRL devices only
    For Each resourceName In resources
        If InStr(1, resourceName, "ASRL") > 0 Then
            asrlResources = asrlResources & "," & resourceName
        End If
    Next resourceName

    Set RM_P = Nothing

    ' Populate dropdown list
    With commCell.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
             Formula1:=asrlResources
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowInput = True
        .ShowError = True
    End With
End Sub


