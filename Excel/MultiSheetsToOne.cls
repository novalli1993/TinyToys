VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub MultiSheetsToOne()
Application.ScreenUpdating = False
Application.EnableEvents = False
Dim path As String
Dim wb As Workbook
Dim ws As Worksheet
Dim validRange As Range

Dim firstCell As Range
Dim lastCell As Range
Dim rangeToCopy As Range
Dim continue As Long


path = "J:\ExcelProcess\Ô­Ê¼Êý¾Ý.xlsx"
Set wb = Workbooks.Open(path, ReadOnly:=True)

For i = 1 To wb.Sheets.Count:
    If i Mod 5 Then
        DoEvents
    End If
    Set validRange = wb.Sheets(i).usedRange
    If ThisWorkbook.Sheets(1).usedRange.Count = 1 And IsEmpty(ThisWorkbook.Sheets(1).usedRange.Value) Then
        Set firstCell = validRange.Cells(validRange.Row, validRange.Column + 1)
        continue = 1
    Else:
        Set firstCell = validRange.Cells(validRange.Row + 1, validRange.Column + 1)
        continue = ThisWorkbook.Sheets(1).usedRange.Row + ThisWorkbook.Sheets(1).usedRange.Rows.Count
    End If
    Set lastCell = validRange.Cells(validRange.Row + validRange.Rows.Count - 1, validRange.Column + validRange.Columns.Count - 1)
    Set rangeToCopy = validRange.Range(firstCell, lastCell)
    rangeToCopy.Copy
    ThisWorkbook.Sheets(1).Rows(continue).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
Next i

Application.ScreenUpdating = True
Application.EnableEvents = True
End Sub

