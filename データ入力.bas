Attribute VB_Name = "Module2"
Option Explicit

Sub sheet()

Dim tws As Worksheet
Set tws = ActiveSheet
Dim i As Long
Dim OpenFileName As Workbook

'Set OpenFileName = Application.GetOpenFilename("Excelブック,*.excel,CSVファイル,*.csv")


Workbooks.Open ThisWorkbook.Path & "\201005客先タイムシート.csv"

Dim ts As Workbook
Set ts = Workbooks("201005客先タイムシート.csv")

last = Cells(Rows.Count, 1).End(xlUp).Row

    With ts.Worksheets("201005客先タイムシート")
    
    For i = 2 To last

        tws.Cells(i + 1, 1).Value = .Cells(i, 2).Value
        tws.Cells(i + 1, 2).Value = .Cells(i, 3).Value
        tws.Cells(i + 1, 3) = Application.WorksheetFunction.Sum(Range(.Cells(i, 6), .Cells(i, 11)))
 
    Next

End With


ts.Close

'-----socia-----

Dim result As Double: result = 0

Workbooks.Open ThisWorkbook.Path & "\202005内部タイムシート.csv"

Dim sca As Workbook
Set sca = Workbooks("202005内部タイムシート.csv")

last = Cells(Rows.Count, 1).End(xlUp).Row

With sca.Worksheets("202005内部タイムシート")

    For i = 2 To last

    On Error Resume Next

    result = WorksheetFunction.Match(.Cells(i, 1), Range(tws.Cells(3, 1), tws.Cells(149, 1)), 0)

    On Error GoTo 0

        If result <> 0 Then
    
            tws.Cells(result + 2, 4).Value = .Cells(i, 1).Value
            tws.Cells(result + 2, 5).Value = .Cells(i, 2).Value
            tws.Cells(result + 2, 6).Value = .Cells(i, 7) - Application.WorksheetFunction.Sum(Range(.Cells(i, 8), .Cells(i, 10)))

        result = 0

    End If

Next

    tws.Columns("F").NumberFormat = "[h]:mm"

End With

sca.Close

End Sub


