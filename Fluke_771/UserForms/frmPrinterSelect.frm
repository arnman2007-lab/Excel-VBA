VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmPrinterSelect 
   Caption         =   "UserForm1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "frmPrinterSelect.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmPrinterSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Click()

End Sub
Private Sub UserForm_Initialize()
    Dim objWMI As Object
    Dim objPrinters As Object
    Dim objPrinter As Object

    Set objWMI = GetObject("winmgmts:\\.\root\cimv2")
    Set objPrinters = objWMI.ExecQuery("Select * From Win32_Printer")

    For Each objPrinter In objPrinters
        cmbPrinters.AddItem objPrinter.Name
    Next objPrinter
End Sub

Private Sub cmdPrint_Click()
    Dim selectedPrinter As String
    Dim ws As Worksheet

    selectedPrinter = cmbPrinters.Value
    If selectedPrinter = "" Then
        MsgBox "Please select a printer.", vbExclamation
        Exit Sub
    End If

    On Error GoTo PrinterError
    Dim p As String
p = GetPrinterConnectionString(selectedPrinter)

If p = "" Then
    MsgBox "Could not match printer with a valid Excel connection string.", vbCritical
    Exit Sub
Else
    Application.ActivePrinter = p
End If
    On Error GoTo 0

    Set ws = ThisWorkbook.Sheets("LabelSheet")
    If ws.PageSetup.PrintArea <> "" Then
        ws.PrintOut
    Else
        MsgBox "Print area not defined on LabelSheet.", vbExclamation
    End If

    Unload Me
    Exit Sub

PrinterError:
    MsgBox "Could not set printer: " & selectedPrinter, vbCritical
End Sub

Function GetPrinterConnectionString(printerName As String) As String
    Dim prn As String
    Dim i As Integer
    On Error Resume Next

    For i = 0 To 99
        prn = Application.Printers(i)
        If InStr(1, prn, printerName, vbTextCompare) > 0 Then
            GetPrinterConnectionString = prn
            Exit Function
        End If
    Next i

    GetPrinterConnectionString = ""
End Function

