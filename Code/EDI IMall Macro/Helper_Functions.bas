Attribute VB_Name = "Helper_Functions"
Option Explicit

'List of custom error messages
Enum CustErr
    COLNOTFOUND = 50000
    VALIDATIONFAILED = 50001
End Enum

'---------------------------------------------------------------------------------------
' Proc : UserImportFile
' Date : 1/29/2013
' Desc : Prompts the user to select a file for import
'---------------------------------------------------------------------------------------
Sub UserImportFile(DestRange As Range, Optional DelFile As Boolean = False, Optional ShowAllData As Boolean = False, Optional SourceSheet As String = "")
    Dim File As String              'Full path to user selected file
    Dim FileDate As String          'Date the file was last modified
    Dim OldDispAlert As Boolean     'Original state of Application.DisplayAlerts

    OldDispAlert = Application.DisplayAlerts
    File = Application.GetOpenFilename()

    Application.DisplayAlerts = False
    If File <> "False" Then
        FileDate = Format(FileDateTime(File), "mm/dd/yy")
        Workbooks.Open File
        If SourceSheet = "" Then SourceSheet = ActiveSheet.Name
        If ShowAllData = True Then
            On Error Resume Next
            ActiveSheet.AutoFilter.ShowAllData
            ActiveSheet.UsedRange.Columns.Hidden = False
            ActiveSheet.UsedRange.Rows.Hidden = False
            On Error GoTo 0
        End If
        Sheets(SourceSheet).UsedRange.Copy Destination:=DestRange
        ActiveWorkbook.Close
        ThisWorkbook.Activate

        If DelFile = True Then
            DeleteFile File
        End If
    Else
        Err.Raise 18
    End If
    Application.DisplayAlerts = OldDispAlert
End Sub

'---------------------------------------------------------------------------------------
' Proc : GetWorkbookPath
' Date : 3/19/2013
' Desc : Gets the full path of ThisWorkbook
'---------------------------------------------------------------------------------------
Function GetWorkbookPath() As String
    Dim fullName As String
    Dim wrkbookName As String
    Dim pos As Long

    wrkbookName = ThisWorkbook.Name
    fullName = ThisWorkbook.fullName

    pos = InStr(1, fullName, wrkbookName, vbTextCompare)

    GetWorkbookPath = Left$(fullName, pos - 1)
End Function

'---------------------------------------------------------------------------------------
' Proc : EndsWith
' Date : 3/19/2013
' Desc : Checks if a string ends in a specified character
'---------------------------------------------------------------------------------------
Function EndsWith(ByVal InString As String, ByVal TestString As String) As Boolean
    EndsWith = (Right$(InString, Len(TestString)) = TestString)
End Function

'---------------------------------------------------------------------------------------
' Proc : DeleteColumn
' Date : 4/11/2013
' Desc : Removes a column based on text in the column header
'---------------------------------------------------------------------------------------
Sub DeleteColumn(HeaderText As String)
    Dim i As Integer

    For i = ActiveSheet.UsedRange.Columns.Count To 1 Step -1
        If Trim(Cells(1, i).Value) = HeaderText Then
            Columns(i).Delete
            Exit For
        End If
    Next
End Sub

'---------------------------------------------------------------------------------------
' Proc : FindColumn
' Date : 4/11/2013
' Desc : Returns the column number if a match is found
'---------------------------------------------------------------------------------------
Function FindColumn(ByVal HeaderText As String, Optional SearchArea As Range) As Integer
    Dim i As Integer: i = 0
    Dim ColText As String

    If TypeName(SearchArea) = "Nothing" Or TypeName(SearchArea) = Empty Then
        Set SearchArea = Range(Cells(1, 1), Cells(1, ActiveSheet.UsedRange.Columns.Count))
    End If

    For i = 1 To SearchArea.Columns.Count
        ColText = Trim(SearchArea.Cells(1, i).Value)

        Do While InStr(ColText, "  ")
            ColText = Replace(ColText, "  ", " ")
        Loop

        If ColText = HeaderText Then
            FindColumn = i
            Exit For
        End If
    Next

    If FindColumn = 0 Then Err.Raise CustErr.COLNOTFOUND, "FindColumn", HeaderText
End Function
