Attribute VB_Name = "SelectingPrinters"
Sub SelectAndPrint()
    Dim bChoice As Boolean
    Dim printerName As String
    Dim ws As Worksheet
    Dim wasHidden As Boolean

    Set ws = ThisWorkbook.Sheets("LabelSheet")

    ' Show printer setup dialog
    bChoice = Application.Dialogs(xlDialogPrinterSetup).show(Application.ActivePrinter)

    If Not bChoice Then
        MsgBox "Printer selection cancelled by user.", vbInformation
        Exit Sub
    End If

    printerName = Application.ActivePrinter

    ' Check that a print area is defined
    If ws.PageSetup.PrintArea = "" Then
        MsgBox "No print area defined on LabelSheet.", vbExclamation
        Exit Sub
    End If

    ' Temporarily unhide if needed
    wasHidden = (ws.Visible <> xlSheetVisible)
    If wasHidden Then ws.Visible = xlSheetVisible

    ' Print the sheet
    ws.PrintOut ActivePrinter:=printerName, Copies:=1, Collate:=True

    ' Re-hide if it was hidden before
    If wasHidden Then ws.Visible = xlSheetHidden
End Sub

