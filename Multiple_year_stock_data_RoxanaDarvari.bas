Attribute VB_Name = "Module1"
Sub AnalyzeStockSheet()

    Dim wb As Workbook
    Set wb = ThisWorkbook
    
    Dim ws As Worksheet
    
    For Each ws In wb.Sheets
  
    
    Dim lastRow As Long
    Dim openStart As Double
    Dim totalVolume As Double
    Dim ticker As String
    Dim tableRow As Integer
    Dim closeEnd As Double
    Dim yearlyChange As Double
    Dim percentChange As Double
    Dim i As Long
    
    Dim greatestIncrease As Double
    Dim greatestDecrease As Double
    Dim greatestVolume As Double
    Dim tickerGreatestIncrease As String
    Dim tickerGreatestDecrease As String
    Dim tickerGreatestVolume As String
    
    tableRow = 2
    totalVolume = 0
    
    greatestDecrease = 0
    greatestIncrease = 0
    greatestVolume = 0
    
    tickerGreatestIncrease = ""
    tickerGreatestDecrease = ""
    tickerGreatestVolume = ""
    
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    
   
    openStart = ws.Cells(2, 3).Value
    
    For i = 2 To lastRow
    
        totalVolume = totalVolume + ws.Cells(i, 7).Value
        
       
        If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
            
            closeEnd = ws.Cells(i, 6).Value
            
           
            yearlyChange = closeEnd - openStart
            If yearlyChange < 0 Then
                    ws.Cells(tableRow, 10).Interior.Color = 255 ' Red
                ElseIf yearlyChange > 0 Then
                    ws.Cells(tableRow, 10).Interior.Color = RGB(0, 255, 0) ' Green
                Else
                    ' No color if Yearly Change is 0
                    ws.Cells(tableRow, 10).Interior.ColorIndex = xlNone
                End If
                
            ticker = ws.Cells(i, 1).Value
            
            If openStart <> 0 And yearlyChange <> 0 Then
                percentChange = (yearlyChange / openStart) * 100
                
                If percentChange > greatestIncrease Then
                    greatestIncrease = percentChange
                    tickerGreatestIncrease = ticker
                End If
                
                If percentChange < greatestDecrease Then
                    greatestDecrease = percentChange
                    tickerGreatestDecrease = ticker
                End If
                
                If totalVolume > greatestVolume Then
                    greatestVolume = totalVolume
                    tickerGreatestVolume = ticker
                    
                End If


            Else
                percentChange = 0
            End If
            
            
            
            
            ticker = ws.Cells(i, 1).Value
            
           
            ws.Cells(tableRow, 9).Value = ticker
            ws.Cells(tableRow, 10).Value = yearlyChange
            ws.Cells(tableRow, 11).Value = percentChange
            ws.Cells(tableRow, 12).Value = totalVolume
            
            
            tableRow = tableRow + 1
            totalVolume = 0
            
            
            If i + 1 <= lastRow Then
                openStart = ws.Cells(i + 1, 3).Value
            End If
        End If
    Next i
    
    
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(2, 16).Value = tickerGreatestIncrease
    ws.Cells(2, 17).Value = greatestIncrease

    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(3, 16).Value = tickerGreatestDecrease
    ws.Cells(3, 17).Value = greatestDecrease

    ws.Cells(4, 15).Value = "Greatest Total Volume"
    ws.Cells(4, 16).Value = tickerGreatestVolume
    ws.Cells(4, 17).Value = greatestVolume
    
    
    ws.Columns("J:J").NumberFormat = "0.00"
    ws.Columns("K:K").NumberFormat = "0.00"
    ws.Cells(4, 17).NumberFormat = "0.00E+00"
    
    ws.Columns("L:L").NumberFormat = "#,##0"
    
    ws.Columns("J:L").AutoFit
    ws.Columns("O:Q").AutoFit
    
    Next ws
    
    
    MsgBox "Stock analysis for all tickers is complete."
End Sub
