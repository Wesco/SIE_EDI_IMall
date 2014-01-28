Attribute VB_Name = "CreateReport"
Option Explicit

Sub CreateEDI(PO As String, Branch As String, DPC As String)
    Dim vCell As Range
    Dim iRows As Long
    Dim ColQty As Long
    Dim ColPrice As Long
    Dim ColUPC As Long
    Dim ColItmNum As Long
    Dim ColDesc As Long

    Sheets("Cart").Select
    iRows = ActiveSheet.UsedRange.Rows.Count

    'Get the column numbers of columns needed for EDI
    On Error GoTo Col_Not_Found
    ColQty = FindColumn("Quantity")
    ColPrice = FindColumn("Customer Price (USD)")
    ColUPC = FindColumn("UPC")
    ColItmNum = FindColumn("Item Number")
    ColDesc = FindColumn("Description")
    On Error GoTo 0

    'Convert UPC to SIM
    For Each vCell In Range(Cells(2, ColUPC), Cells(iRows, ColUPC))
        vCell.NumberFormat = "@"
        vCell.Value = Left(Right(vCell.Text, 12), 11)
    Next

    'Remove commas, periods, and slashes from item descriptions
    With Range(Cells(2, ColDesc), Cells(iRows, ColDesc))
        .Replace What:=",", Replacement:=" ", LookAt:=xlPart, SearchOrder:=xlByRows
        .Replace What:=".", Replacement:="", LookAt:=xlPart, SearchOrder:=xlByRows
        .Replace What:="\", Replacement:=" ", LookAt:=xlPart, SearchOrder:=xlByRows
        .Replace What:="/", Replacement:=" ", LookAt:=xlPart, SearchOrder:=xlByRows
    End With


    'EDI Data Layout
    'PO_NUMBER, BRANCH, DPC, CUST_LINE, QTY, UOM, UNIT_PRICE, SIM, PART_NO, DESC, SHIP_DATE, SHIPTO, NOTE1, NOTE2
    '    A        B      C       D       E    F       G        H       I      J       K         L      M      N
    '    1        2      3       4       5    6       7        8       9      10      11        12     13     14


    'Fill in EDI data that does com from the cart
    Sheets("EDI").Select
    'EDI Col A = Customer PO
    Range(Cells(1, 1), Cells(iRows - 1, 1)).NumberFormat = "@"
    Range(Cells(1, 1), Cells(iRows - 1, 1)).Value = PO
    'EDI Col B = Branch Number
    Range(Cells(1, 2), Cells(iRows - 1, 2)).NumberFormat = "@"
    Range(Cells(1, 2), Cells(iRows - 1, 2)).Value = Branch
    'EDI Col C = Customer DPC
    Range(Cells(1, 3), Cells(iRows - 1, 3)).NumberFormat = "@"
    Range(Cells(1, 3), Cells(iRows - 1, 3)).Value = DPC
    'EDI Col F = UOM
    Range(Cells(1, 6), Cells(iRows - 1, 6)).Value = "E"
    'EDI Col G = Unit Price
    Range(Cells(1, 7), Cells(iRows - 1, 7)).Value = 0

    'Copy data from cart to EDI sheet
    Sheets("Cart").Select
    'EDI Col E = Quantity
    Range(Cells(2, ColQty), Cells(iRows, ColQty)).Copy
    Sheets("EDI").Range("E1").PasteSpecial xlPasteValuesAndNumberFormats
    'EDI Col H = SIM number
    Range(Cells(2, ColUPC), Cells(iRows, ColUPC)).Copy
    Sheets("EDI").Range("H1").PasteSpecial xlPasteValuesAndNumberFormats
    'EDI Col I = Part Number
    Range(Cells(2, ColItmNum), Cells(iRows, ColItmNum)).Copy
    Sheets("EDI").Range("I1").PasteSpecial xlPasteValuesAndNumberFormats
    'EDI Col M = Note1
    Range(Cells(2, ColDesc), Cells(iRows, ColDesc)).Copy
    Sheets("EDI").Range("M1").PasteSpecial xlPasteValuesAndNumberFormats

    Exit Sub


Col_Not_Found:
    MsgBox "Error - Column '" & Err.Description & "' was not found."
    Exit Sub

End Sub
