Attribute VB_Name = "Module2"
Sub AllPagesWithGreatest()

Dim ws As Worksheet

For Each ws In Worksheets

    Dim Ticker As String
    Dim Volume As Double
    Dim OpenYear As Double
    Dim CloseYear As Double
    Dim RowSum As Integer
    RowSum = 2

    ws.Cells(1, 10).Value = "Ticker"
    ws.Cells(1, 11).Value = "Yearly Change"
    ws.Cells(1, 12).Value = "Percent Change"
    ws.Cells(1, 13).Value = "Total Stock Volume"

    Dim lastRow As Double
    lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    For i = 2 To lastRow
        If (ws.Cells(i, 3).Value = 0) Then
            If (ws.Cells(i + 1).Value <> ws.Cells(i, 1).Value) Then
                Ticker = ws.Cells(i, 1).Value
            End If
        ElseIf (ws.Cells(i + 1, 1).Value = ws.Cells(i, 1).Value) Then
            Volume = Volume + ws.Cells(i, 7).Value
            If (ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value) Then
                OpenYear = ws.Cells(i, 3).Value
            End If
        Else
            Ticker = ws.Cells(i, 1).Value
            Volume = Volume + ws.Cells(i, 7).Value
            CloseYear = ws.Cells(i, 6).Value
            ws.Cells(RowSum, 10).Value = Ticker
            ws.Cells(RowSum, 13).Value = Volume
            If (Volume > 0) Then
                ws.Cells(RowSum, 11).Value = CloseYear - OpenYear
                    If (ws.Cells(RowSum, 11).Value > 0) Then
                        ws.Cells(RowSum, 11).Interior.ColorIndex = 4
                    Else
                        ws.Cells(RowSum, 11).Interior.ColorIndex = 3
                    End If
                ws.Cells(RowSum, 12).Value = ws.Cells(RowSum, 11).Value / OpenYear
            Else
                ws.Cells(RowSum, 11).Value = 0
                ws.Cells(RowSum, 12).Value = 0
            End If
            ws.Cells(RowSum, 12).Style = "percent"
            Volume = 0
            RowSum = RowSum + 1
        End If
    Next i

    Dim TotVol As Double

        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"

    TotVol = 0

    RowSum = RowSum - 2

    For i = 2 To RowSum
        If (ws.Cells(i, 13).Value > TotVol) Then
            TotVol = ws.Cells(i, 13).Value
            ws.Cells(4, 16).Value = ws.Cells(i, 10).Value
        End If
    Next i
    
    ws.Cells(4, 17).Value = TotVol
    
    Dim PctInc As Double
    Dim PctDec As Double
    PctInc = 0
    PctDec = 0

    For i = 2 To RowSum
        If (ws.Cells(i, 12).Value > PctInc) Then
            PctInc = ws.Cells(i, 12).Value
            ws.Cells(2, 16) = ws.Cells(i, 10).Value
        ElseIf (ws.Cells(i, 12).Value < PctDec) Then
            PctDec = ws.Cells(i, 12).Value
            ws.Cells(3, 16).Value = ws.Cells(i, 10).Value
        End If
    Next i
    
    ws.Cells(2, 17).Value = PctInc
    ws.Cells(3, 17).Value = PctDec
    ws.Cells(2, 17).Style = "percent"
    ws.Cells(3, 17).Style = "percent"

    ws.Columns("J:Q").AutoFit

Next ws

End Sub

