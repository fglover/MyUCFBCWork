Attribute VB_Name = "Module1"
Sub Easy()

Dim ticker As String

Dim ticker_total As Double

Dim ws As Worksheet

Dim last_row As String

Dim Summary_Table_Row As Integer

Summary_Table = 2

For Each ws In Worksheets

    ticker_total = 0

    k = 0

last_row = Cells(Rows.Count, "A").End(xlUp).Row

For i = 2 To last_row


If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
    ticker = ws.Cells(i, 1).Value
    
    ticker_total = ticker_total + ws.Cells(i, 7).Value
    
    ws.Range("J" & 2 + k).Value = ticker
    
    ws.Range("K" & 2 + k).Value = ticker_total
    
    ticker_total = 0
 
    k = k + 1

Else

 ticker_total = ticker_total + ws.Cells(i, 7).Value
 
End If

Next i
    ticker_total = 0
 
    k = 0
    
Next ws

End Sub

