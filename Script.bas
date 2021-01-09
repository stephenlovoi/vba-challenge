Attribute VB_Name = "Module1"
Sub Testing()
    
    Dim lastRow, lastColumn As Long
        
    Dim tickerName As String
    Dim annualOpen As Double
    Dim annualClose As Double
    Dim yearlyChange As Double
    Dim volume As Double
    Dim percentChange As Double
    Dim flg As Boolean
    Dim ws As Worksheet
    Dim summaryTableRow As Integer

    For Each ws In Worksheets
    
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        ws.Range("I:L").ClearContents
        lastColumn = ws.Cells(1, Columns.Count).End(xlToLeft).Column
        
        ws.Range("I1") = "Ticker"
        ws.Range("J1") = "Yearly Change"
        ws.Range("K1") = "Percent Change"
        ws.Range("L1") = "Total Stock Volume"
          
        summaryTableRow = 2
        volume = 0
        
        flg = True
        
        For i = 2 To lastRow
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                tickerName = ws.Cells(i, 1).Value
                volume = volume + ws.Cells(i, 7).Value
                ws.Range("I" & summaryTableRow).Value = tickerName
                ws.Range("L" & summaryTableRow).Value = volume
                annualClose = ws.Cells(i, 6).Value
                ws.Range("J" & summaryTableRow).Value = annualClose - annualOpen
                If annualClose - annualOpen > 0 Then
                    ws.Range("J" & summaryTableRow).Interior.Color = vbGreen
                Else
                    ws.Range("J" & summaryTableRow).Interior.Color = vbRed
                End If
                If annualClose = 0 Then
                    If annualOpen = 0 Then
                        ws.Range("K" & summaryTableRow).Value = FormatPercent(0, 2)
                    Else
                        ws.Range("K" & summaryTableRow).Value = FormatPercent(-1, 2)
                    End If
                    Else
                    ws.Range("K" & summaryTableRow).Value = FormatPercent((annualClose - annualOpen) / annualClose, 2)
                End If
                summaryTableRow = summaryTableRow + 1
                volume = 0
                flg = True
            Else
                volume = volume + ws.Cells(i, 7).Value
                If flg = True Then
                    annualOpen = ws.Cells(i, 3).Value
                    flg = False
                End If
            End If
    
        Next i

    Next ws

End Sub
