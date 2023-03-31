Attribute VB_Name = "Module1"
Sub execute()
    Dim xSh As Worksheet
    Application.ScreenUpdating = False
    For Each xSh In Worksheets
        xSh.Select
        Call stocks
        Call table
    Next
    Application.ScreenUpdating = True
End Sub


Sub stocks()
    Dim tickerRow As Integer
    Dim yearlyChange, percentChange, changeOpen, changeClose As Double
    Dim lastRow As LongPtr
    Dim stockTotal As LongPtr
    Dim tickerSymbol As String
    
    tickerRow = 2
    stockTotal = 0
    yearlyChange = 0
    percentChange = 0
    
    lastRow = Range("A" & Rows.Count).End(xlUp).Row

    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
    ActiveSheet.Columns("I:M").AutoFit
    Range("K1").EntireColumn.NumberFormat = "0.00%"
    
    changeOpen = Cells(2, 3).Value
    
    For i = 2 To lastRow
        
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            
            tickerSymbol = Cells(i, 1).Value
            
            stockTotal = stockTotal + Cells(i, 7).Value
            
            changeClose = Cells(i, 6).Value
            
            yearlyChange = changeClose - changeOpen
            percentChange = (changeClose - changeOpen) / changeOpen
            
            Range("I" & tickerRow).Value = tickerSymbol
            Range("J" & tickerRow).Value = yearlyChange
            Range("K" & tickerRow).Value = percentChange
            Range("L" & tickerRow).Value = stockTotal
            
            tickerRow = tickerRow + 1
            
            If Cells(i, 2).Value > Cells(i + 1, 2).Value Then
            
                changeOpen = Cells(i + 1, 3).Value
                
            End If
            
            stockTotal = 0
            
        ElseIf Cells(i, 2).Value > Cells(i + 1, 2).Value Then
            
            changeOpen = Cells(i, 3).Value
            
        Else
            
            stockTotal = stockTotal + Cells(i, 7).Value
        
        End If
        
        
        
        
    Next i
    For j = 2 To lastRow
        If Cells(j, 10).Value > 0 Then
            Cells(j, 10).Interior.ColorIndex = 4
        ElseIf Cells(j, 10).Value < 0 Then
            Cells(j, 10).Interior.ColorIndex = 3
        End If
    Next j
    

End Sub

Sub table()
    Dim tickerSymbol As String
    
    tableLastRow = Range("I" & Rows.Count).End(xlUp).Row
    
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
    ActiveSheet.Columns("O:Q").AutoFit
    
    startNumber = Range("K2").Value
    
    For i = 2 To tableLastRow
    
        If Cells(i, 11).Value > startNumber Then
            
            startNumber = Cells(i, 11).Value
            tickerSymbolI = Cells(i, 9).Value
            increaseNumber = startNumber
            
        End If
        
    Next i
    
    For k = 2 To tableLastRow
        If Cells(k, 11).Value < startNumber Then
            
            startNumber = Cells(k, 11).Value
            tickerSymbolD = Cells(k, 9).Value
            decreaseNumber = startNumber
            
        End If
        
    Next k
    
    For j = 2 To tableLastRow
    
        If Cells(j, 12).Value > stockNumber Then
            
            stockNumber = Cells(j, 12).Value
            tickerSymbolH = Cells(j, 9).Value
            highestNumber = stockNumber
            
        End If
        
    Next j
    
    Cells(2, 17).Value = increaseNumber
    Cells(3, 17).Value = decreaseNumber
    Cells(4, 17).Value = highestNumber
    Cells(2, 16).Value = tickerSymbolI
    Cells(3, 16).Value = tickerSymbolD
    Cells(4, 16).Value = tickerSymbolH
    Range("Q1").EntireColumn.NumberFormat = "0.00%"
    
End Sub


