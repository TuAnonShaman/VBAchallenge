# VBA Challenge Code - Stephen Grantham

Sub StockTest()



'Modify to run on all worksheets

For Each ws In Worksheets

    ws.Range("J1").Value = "Ticker"
    ws.Range("K1").Value = "Yearly Change"
    ws.Range("L1").Value = "Percent Change"
    ws.Range("M1").Value = "Total Stock Volume"
    
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
        
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"




'Find & label last row

    Dim LastRow As LongLong
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    
    
    
'Grab all tickers

    Dim Ticker As String
        
    Dim SummaryTableRow As Integer
        SummaryTableRow = 2
    
    For i = 2 To LastRow
        
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            Ticker = ws.Cells(i, 1).Value
            
            ws.Range("J" & SummaryTableRow).Value = Ticker
            
            SummaryTableRow = SummaryTableRow + 1
            
        End If
        
    Next i



'Grab all yearly & percent changes

    SummaryTableRow = 2

    Dim StockOpen As Double
    Dim StockClose As Double
    Dim YearChange As Double
    Dim PercentChange As Double


    For i = 2 To LastRow
        
        If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
            StockOpen = ws.Cells(i, 3).Value

        ElseIf ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            StockClose = ws.Cells(i, 6).Value
            
            YearChange = StockClose - StockOpen
            PercentChange = (StockClose - StockOpen) / StockOpen
            
            ws.Range("K" & SummaryTableRow).Value = YearChange
            ws.Range("L" & SummaryTableRow).Value = PercentChange
                ws.Range("L" & SummaryTableRow).NumberFormat = "0.00%"
            
            SummaryTableRow = SummaryTableRow + 1
   
         End If
    
    Next i


'Conditional Format Percent Change
    
    Dim SummaryLastRow As Long
    
    SummaryLastRow = ws.Cells(Rows.Count, 12).End(xlUp).Row
    
    
    For i = 2 To SummaryLastRow
    
        If ws.Cells(i, 11) > 0 Then
            ws.Cells(i, 11).Interior.ColorIndex = 4
            
        ElseIf ws.Cells(i, 11) < 0 Then
            ws.Cells(i, 11).Interior.ColorIndex = 3
                        
        ElseIf ws.Cells(i, 11) = 0 Then
            ws.Cells(i, 11).Interior.ColorIndex = None
        
        End If
    
    Next i
    
    

'Calculate Stock Volume

    SummaryTableRow = 2
    
    Dim StockVolume As LongLong
        StockVolume = 0

    For i = 2 To LastRow
    
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1) Then
            StockVolume = StockVolume + ws.Cells(i, 7).Value
            
            ws.Range("M" & SummaryTableRow).Value = StockVolume
            
            StockVolume = 0
            SummaryTableRow = SummaryTableRow + 1
            
        Else
        
            StockVolume = StockVolume + ws.Cells(i, 7).Value
        
        End If
        
    Next i



'Grab Greatest % Increase

    Dim GincStock As String
    Dim GincValue As Double
        GincValue = -1000
    
    For i = 2 To LastRow

        If ws.Cells(i, 12).Value > GincValue Then
            GincStock = ws.Cells(i, 10).Value
            GincValue = ws.Cells(i, 12).Value
            
        End If
    
    Next i
    
    ws.Range("P2") = GincStock
        ws.Range("P2").NumberFormat = "0.00%"
    ws.Range("Q2") = GincValue
        ws.Range("Q2").NumberFormat = "0.00%"



'Grab Greatest % Decrease

    Dim GdecStock As String
    Dim GdecValue As Double
        GdecValue = 1000
    
    For i = 2 To LastRow

        If ws.Cells(i, 12).Value < GdecValue Then
            GdecStock = ws.Cells(i, 10).Value
            GdecValue = ws.Cells(i, 12).Value
            
        End If
    
    Next i
    
    ws.Range("P3") = GdecStock
        ws.Range("P3").NumberFormat = "0.00%"
    ws.Range("Q3") = GdecValue
        ws.Range("Q3").NumberFormat = "0.00%"



'Grab Greatest Total Volume

    Dim GvolStock As String
    Dim GvolValue As LongLong
        GvolValue = 0
    
    For i = 2 To LastRow

        If ws.Cells(i, 13).Value > GvolValue Then
            GvolStock = ws.Cells(i, 10).Value
            GvolValue = ws.Cells(i, 13).Value
            
        End If
    
    Next i
    
    ws.Range("P4") = GvolStock
    ws.Range("Q4") = GvolValue


'AutoFit Columns

    ws.Range("J1:Q1").Font.Bold = True
    ws.Range("J:Q").Columns.AutoFit


Next ws

End Sub

