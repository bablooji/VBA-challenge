Attribute VB_Name = "Module1"
'Module 2 Challenge - Bhagya Prasad
'Stock Symbol Summary

Sub getTickerDetails()

    For Each ws In Worksheets
     Dim row1, col1 As Integer
            Dim WorksheetName As String
            lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
            WorksheetName = ws.Name
            ws.Range("H1").EntireColumn.Insert
            ws.Cells(1, 9).Value = WorksheetName                     'I
            ws.Cells(1, 10).Value = "Ticker"                                 'J
            ws.Cells(1, 11).Value = "Yearly Change"                  'K
            ws.Cells(1, 12).Value = "Percent Change"                 'L
            ws.Cells(1, 13).Value = "Total Stock Volume"          'M
            ws.Cells(2, 16).Value = "Greatest % Increase"           'N2
            ws.Cells(3, 16).Value = "Greatest % Decrease"         'N3
            ws.Cells(4, 16).Value = "Greatest Total Volume"       'N4
            ws.Cells(1, 17).Value = "Ticker"                                'O1
            ws.Cells(1, 18).Value = "Value"                                 'P1
            'ws.Range("H2:H" & lastrow) = ws.Cells(2, 1).Value
        
            Dim tickerSymbol As String
            Dim tickerVol As Double
            tickerVol = 0

        'For each ticker name generate the summary table
        Dim uniqueTickerRow As Integer
        uniqueTickerRow = 2
        Dim openingPrice As Double
        Dim closingPrice As Double
        Dim yearlyChange As Double
        Dim percentChange As Double

        
       For row1 = 2 To lastrow
            If ws.Cells(row1 + 1, 1).Value <> ws.Cells(row1, 1).Value Then
                tickerSymbol = ws.Cells(row1, 1).Value
                tickerVol = tickerVol + ws.Cells(row1, 7).Value
                ws.Range("J" & uniqueTickerRow).Value = tickerSymbol
                ws.Range("M" & uniqueTickerRow).Value = tickerVol
                closingPrice = ws.Cells(row1, 6).Value
                yearlyChange = (closingPrice - openingPrice)
                ws.Range("K" & uniqueTickerRow).Value = yearlyChange
                
                If openingPrice = 0 Then
                    percentChange = 0
                Else
                    percentChange = yearlyChange / openingPrice
            End If

                ws.Range("L" & uniqueTickerRow).Value = percentChange
                ws.Range("L" & uniqueTickerRow).NumberFormat = "0.00%"
                uniqueTickerRow = uniqueTickerRow + 1
                tickerVol = 0
                openingPrice = ws.Cells(row1 + 1, 3)
            
            Else
                tickerVol = tickerVol + ws.Cells(row1, 7).Value
            End If
        
        Next row1
        
'Yearly Change conditional formatting using color coding
        
        outPutTable = ws.Cells(Rows.Count, 11).End(xlUp).Row
    
        For row1 = 2 To outPutTable
            If ws.Cells(row1, 11).Value > 0 Then
                ws.Cells(row1, 11).Interior.ColorIndex = 10
            Else
                ws.Cells(row1, 11).Interior.ColorIndex = 3
            End If
        Next row1

'Populate the greatest changes table
               For row1 = 2 To outPutTable
            'Find the maximum percent change
            If ws.Cells(row1, 12).Value = Application.WorksheetFunction.Max(ws.Range("L2:L" & outPutTable)) Then
                ws.Cells(2, 17).Value = ws.Cells(row1, 10).Value
                ws.Cells(2, 18).Value = ws.Cells(row1, 12).Value
                ws.Cells(2, 18).NumberFormat = "0.00%"

            'Find the minimum percent change
            ElseIf ws.Cells(row1, 12).Value = Application.WorksheetFunction.Min(ws.Range("L2:L" & outPutTable)) Then
                ws.Cells(3, 17).Value = ws.Cells(row1, 10).Value
                ws.Cells(3, 18).Value = ws.Cells(row1, 12).Value
                ws.Cells(3, 18).NumberFormat = "0.00%"
            
            'Find the maximum volume of trade
            ElseIf ws.Cells(row1, 13).Value = Application.WorksheetFunction.Max(ws.Range("M2:M" & outPutTable)) Then
                ws.Cells(4, 17).Value = ws.Cells(row1, 10).Value
                ws.Cells(4, 18).Value = ws.Cells(row1, 13).Value
            
            End If
        
        Next row1
        
    Next ws
End Sub

