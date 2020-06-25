Attribute VB_Name = "Module1"
Option Explicit

Function checkData(i As Long) As Boolean

Dim msg As String
Dim color As Integer
Dim isError As Boolean

color = 6
msg = ""
isError = False



'CまたはF列に値が入っていない場合エラーを表示する
If Cells(i, 3) = "" Then
    msg = "客先タイムシートに時間が入力されていません"
    isError = True

ElseIf Cells(i, 6) = "" Then
    msg = "内部タイムシートに時間が入力されていません"
    isError = True

ElseIf Cells(i, 1) <> Cells(i, 4) Then
    msg = "社員番号が一致していません"
    isError = True

ElseIf Cells(i, 2) <> Cells(i, 5) Then
    msg = "氏名が一致していません"
    isError = True
    
End If


If isError Then
Cells(i, 8).Value = msg
Cells(i, 8).Interior.ColorIndex = color
End If

checkData = isError

End Function

Function checkTime(i As Long)

'〇×を表示する
    If Cells(i, 3) = Cells(i, 7) Then
        Cells(i, 8).Value = "〇"
    Else
        Cells(i, 8).Value = "×"
    End If

    'With Cells(i, 8)
    
        '.Value = "〇"
   ' End With

End Function

Sub check()

Dim i As Long
Dim check As Range
Dim ErrorCount As Long: ErrorCount = 0
Dim a As Worksheet
Set a = ActiveSheet

Dim last As Long
last = a.UsedRange.Rows(a.UsedRange.Rows.Count).Row

'last = Cells(Rows.Count, 1).End(xlUp).row

For i = 3 To last

'Set check = Cells(i, 8)

If checkData(i) Then
    ErrorCount = ErrorCount + 1
    

End If

Next

Columns(8).AutoFit

If ErrorCount > 0 Then

End

End If



For i = 3 To last

'F列を少数に変換する
Cells(i, 7).Value = (Cells(i, 6)) * 24


Call checkTime(i)


Next

Columns(8).AutoFit


'〇と×の数をカウント

Cells(3, 10).Value = WorksheetFunction.CountIf(Range(Cells(3, 8), Cells(last, 8)), "〇")
Cells(3, 11).Value = WorksheetFunction.CountIf(Range(Cells(3, 8), Cells(last, 8)), "×")

Cells(3, 12).Value = Round(WorksheetFunction.Sum(Range(Cells(3, 7), Cells(last, 7))))
Cells(3, 13).Value = Round(WorksheetFunction.Average(Range(Cells(3, 7), Cells(last, 7))))


End Sub



