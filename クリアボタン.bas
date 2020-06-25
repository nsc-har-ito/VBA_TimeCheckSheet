Attribute VB_Name = "Module3"
Option Explicit

Sub clear()

Dim i As Long
Dim last As Long

last = Cells(Rows.Count, 8).End(xlUp).Row
 
i = last

Range(Cells(3, 7), Cells(i, 8)).clear
Columns(8).ColumnWidth = 9

Range(Cells(3, 10), Cells(3, 13)).ClearContents


End Sub

