Attribute VB_Name = "Exports"
Option Explicit

Sub ExportEDI(PO As String)
    Dim PrevDispAlert As Boolean
    Const EDIPath As String = "\\idxexchange-new\EDI\Spreadsheet_PO\"
    Dim FileName As String: FileName = PO & Format(Now, "hhmmss")

    PrevDispAlert = Application.DisplayAlerts
    Application.DisplayAlerts = False

    Sheets("EDI").Copy
    ActiveWorkbook.SaveAs FileName:=EDIPath & FileName, FileFormat:=xlCSV
    ActiveWorkbook.Close

    Application.DisplayAlerts = PrevDispAlert
End Sub
