Attribute VB_Name = "Module2"
Sub SinglePageWithGreatest()

    Dim Ticker As String
    Dim Volume As Double
    Dim OpenYear As Double
    Dim CloseYear As Double
    Dim RowSum As Integer
    RowSum = 2

    Cells(1, 10).Value = "Ticker"
    Cells(1, 11).Value = "Yearly Change"
    Cells(1, 12).Value = "Percent Change"
    Cells(1, 13).Value = "Total Stock Volume"

    Dim lastRow As Double
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row

    For i = 2 To lastRow
        If Cells(i, 3).Value = 0 Then
            If Cells(i + 1).Value <> Cells(i, 1).Value Then
                Ticker = Cells(i, 1).Value
            End If
        ElseIf Cells(i + 1, 1).Value = Cells(i, 1).Value Then
            Volume = Volume + Cells(i, 7).Value
            If Cells(i - 1, 1).Value <> Cells(i, 1).Value Then
                OpenYear = Cells(i, 3).Value
            End If
        Else
            Ticker = Cells(i, 1).Value
            Volume = Volume + Cells(i, 7).Value
            CloseYear = Cells(i, 6).Value
            Cells(RowSum, 10).Value = Ticker
            Cells(RowSum, 13).Value = Volume
            If Volume > 0 Then
                Cells(RowSum, 11).Value = CloseYear - OpenYear
                    If Cells(RowSum, 11).Value > 0 Then
                        Cells(RowSum, 11).Interior.ColorIndex = 4
                    Else
                        Cells(RowSum, 11).Interior.ColorIndex = 3
                    End If
                Cells(RowSum, 12).Value = Cells(RowSum, 11).Value / OpenYear
            Else
                Cells(RowSum, 11).Value = 0
                Cells(RowSum, 12).Value = 0
            End If
            Cells(RowSum, 12).Style = "percent"
            Volume = 0
            RowSum = RowSum + 1
        End If
    Next i

    Dim TotVol As Double

        Cells(2, 15).Value = "Greatest % Increase"
        Cells(3, 15).Value = "Greatest % Decrease"
        Cells(4, 15).Value = "Greatest Total Volume"
        Cells(1, 16).Value = "Ticker"
        Cells(1, 17).Value = "Value"

    TotVol = 0

    RowSum = RowSum - 2

    For i = 2 To RowSum
        If Cells(i, 13).Value > TotVol Then
            TotVol = Cells(i, 13).Value
            Cells(4, 16).Value = Cells(i, 10).Value
        End If
    Next i
    
    Cells(4, 17).Value = TotVol
    
    Dim PctInc As Double
    Dim PctDec As Double
    PctInc = 0
    PctDec = 0

    For i = 2 To RowSum
        If Cells(i, 12).Value > PctInc Then
            PctInc = Cells(i, 12).Value
            Cells(2, 16) = Cells(i, 10).Value
        ElseIf Cells(i, 12).Value < PctDec Then
            PctDec = Cells(i, 12).Value
            Cells(3, 16).Value = Cells(i, 10).Value
        End If
    Next i
    
    Cells(2, 17).Value = PctInc
    Cells(3, 17).Value = PctDec
    Cells(2, 17).Style = "percent"
    Cells(3, 17).Style = "percent"

    Columns("J:Q").AutoFit

End Sub

