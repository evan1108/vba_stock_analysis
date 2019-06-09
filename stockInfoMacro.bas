Attribute VB_Name = "Module1"
Sub stockInfo():
    Dim numRows, j, n, findRow, numUniqueTickers As Long
    Dim c As Integer
    Dim volume, openingPrice, closingPrice, yrlyChange, greatestInc, greatestDec, percentChange, greatestVol As Double
    Dim ticker As String
    
    ' Get the number of rows
    numRows = Cells(Rows.Count, "A").End(xlUp).Row
    
    ' Copy the unique tickers to column K
    c = 2
    For n = 2 To numRows
        If Cells(n, 1).Value <> Cells(n - 1, 1).Value Then
            Range("K" & c).Value = Cells(n, 1).Value
            c = c + 1
        End If
    Next n
    
    For j = 2 To numRows
        ' looks for the first row of a ticker
        If Cells(j, 1).Value = Cells(j + 1, 1).Value And Cells(j, 1) <> Cells(j - 1, 1).Value Then
            openingPrice = Cells(j, 3).Value
            volume = Range("G" & j).Value
        ' gets all the ticker rows in between the first and last row
        ElseIf Cells(j, 1).Value = Cells(j + 1, 1).Value And Cells(j, 1) = Cells(j - 1, 1).Value Then
            volume = volume + Range("G" & j).Value
        ' checks for the last row of a ticker and calculates yearly change, percent change and volume
        ElseIf Cells(j, 1).Value = Cells(j - 1, 1).Value And Cells(j, 1) <> Cells(j + 1, 1).Value And openingPrice <> 0 Then
            closingPrice = Cells(j, 6).Value
            volume = volume + Range("G" & j).Value
            yrlyChange = closingPrice - openingPrice
            percentChange = yrlyChange / openingPrice
            findRow = Cells(Rows.Count, "L").End(xlUp).Row
            Range("L" & findRow + 1).Value = yrlyChange
            Range("M" & findRow + 1).Value = percentChange
            Range("N" & findRow + 1).Value = volume
            volume = 0
        ' catch the corner case where opening price and closing price are 0
        ElseIf Cells(j, 1).Value = Cells(j - 1, 1).Value And Cells(j, 1) <> Cells(j + 1, 1).Value And openingPrice = 0 Then
            closingPrice = Cells(j, 6).Value
            volume = volume + Range("G" & j).Value
            yrlyChange = closingPrice - openingPrice
            percentChange = 0
            findRow = Cells(Rows.Count, "L").End(xlUp).Row
            Range("L" & findRow + 1).Value = yrlyChange
            Range("M" & findRow + 1).Value = percentChange
            Range("N" & findRow + 1).Value = volume
            volume = 0
        End If
    Next j
    
    ' get the number of unique tickers
    numUniqTickers = Cells(Rows.Count, "K").End(xlUp).Row
    
    ' get the greatest increase
    greatestInc = Range("M2").Value
    Range("K2").Value = ticker
    
    For i = 2 To numUniqTickers
        If Range("M" & i).Value > greatestInc Then
            greatestInc = Range("M" & i).Value
            ticker = Range("K" & i).Value
        End If
    Next i
    
    Range("S2").Value = greatestInc
    Range("R2").Value = ticker
    ' get the greastest decrease
    greatestDec = Range("M2").Value
    Range("K2").Value = ticker
    
    For i = 2 To numUniqTickers
        If Range("M" & i).Value < greatestDec Then
            greatestDec = Range("M" & i).Value
            ticker = Range("K" & i).Value
        End If
    Next i
    
    Range("S3").Value = greatestDec
    Range("R3").Value = ticker
    ' get the greatest volume
    greatestVol = Range("N2").Value
    Range("K2").Value = ticker
    
    For i = 2 To numUniqTickers
        If Range("N" & i).Value > greatestVol Then
            greatestVol = Range("N" & i).Value
            ticker = Range("K" & i).Value
        End If
    Next i
    Range("S4").Value = greatestVol
    Range("R4").Value = ticker
    
End Sub



