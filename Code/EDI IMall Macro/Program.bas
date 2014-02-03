Attribute VB_Name = "Program"
Option Explicit
Public Const RepositoryName As String = "SIE_EDI_IMall"
Public Const VersionNumber As String = "1.0.0"

Sub Main()
    Dim PO As String
    Dim DPC As String
    Dim Branch As String

    Clean

    'User import cart file
    On Error GoTo Import_Failed
    UserImportFile Sheets("Cart").Range("A1"), False
    On Error GoTo 0

    'Get DPC from user
    DPC = InputBox("Customer DPC Number:", "Customer DPC")
    If DPC = "" Then
        MsgBox "Order Canceled - a DPC number was not entered."
        GoTo End_Sub
    End If

    'Get PO from user
    PO = InputBox("Customer PO Number:", "Customer PO")
    If PO = "" Then
        MsgBox "Order Canceled - a customer PO was not entered."
        GoTo End_Sub
    End If

    'Get Branch from user
    Branch = InputBox("Branch Number: ", "EDI Branch")
    If Branch = "" Then
        MsgBox "Order Canceled - a branch number was not entered."
        GoTo End_Sub
    End If

    'Create an EDI order out of the cart
    CreateEDI PO, Branch, DPC

    'Save the EDI order
    ExportEDI PO

    Clean
    MsgBox "Order sent!"
    GoTo End_Sub


Import_Failed:
    Clean
    MsgBox "Macro aborted - A Siemens cart was not selected."
    GoTo End_Sub

End_Sub:
    Clean
End Sub

Sub Clean()
    Dim PrevDispAlert As Boolean
    Dim PrevScreUpdat As Boolean

    'Store current settings
    PrevDispAlert = Application.DisplayAlerts
    PrevScreUpdat = Application.ScreenUpdating

    'Disable alerts and screen updating
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    'Remove data
    Sheets("Cart").Cells.Delete
    Sheets("EDI").Cells.Delete

    'Select the macro sheet
    Sheets("Macro").Select

    'Return settings to their previous state
    Application.ScreenUpdating = PrevScreUpdat
    Application.DisplayAlerts = PrevDispAlert
End Sub
