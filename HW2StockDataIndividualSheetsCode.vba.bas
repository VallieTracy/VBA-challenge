Attribute VB_Name = "Module1"
Sub StockData()

    Dim ws As Worksheet
    
    For Each ws In Worksheets
    
        If ws.Name <> "Combined Data" Then
        
            Dim StockSymbol As String
            
            Dim Volume As Double
            Volume = 0
            
            Dim ClosingPrice As Double
            ClosingPrice = 0
            
            Dim OpeningPrice As Double
            OpeningPrice = 0
            
            Dim YearlyChange As Double
            YearlyChange = 0
            YearlyChange = ClosingPrice - OpeningPrice
            
            Dim SummaryTableRow As Integer
            SummaryTableRow = 2
            
            ws.Range("K1").Value = "Ticker"
            ws.Range("L1").Value = "Yearly Change"
            ws.Range("M1").Value = "Percent Change"
            ws.Range("N1").Value = "Volume"
            ws.Range("K1:N1").Font.Bold = True
    
            LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
            
            OpeningPrice = ws.Cells(2, 3).Value
            
             For i = 2 To LastRow
            
                StockSymbol = ws.Cells(i, 1).Value
        
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
                    ws.Range("K" & SummaryTableRow).Value = StockSymbol
                    
                    ClosingPrice = ws.Cells(i, 6).Value
                    
                    YearlyChange = ClosingPrice - OpeningPrice
                    ws.Range("L" & SummaryTableRow).Value = YearlyChange
                    If YearlyChange > 0 Then
                        ws.Range("L" & SummaryTableRow).Interior.ColorIndex = 4
                    ElseIf YearlyChange = 0 Then
                    ws.Range("L" & SummaryTableRow).Interior.ColorIndex = 33
                    ElseIf YearlyChange < 0 Then
                        ws.Range("L" & SummaryTableRow).Interior.ColorIndex = 3
                    End If
                    
                    If OpeningPrice <> 0 Then
                        Dim PercentChange As Double
                        PercentChange = YearlyChange / OpeningPrice
                        ws.Range("M" & SummaryTableRow).Value = PercentChange
                        ws.Range("M" & SummaryTableRow).NumberFormat = "0.00%"
                    Else
                        Dim PercentChangeError As String
                        ws.Range("M" & SummaryTableRow).Value = "n/a"
                        Range("M" & SummaryTableRow).HorizontalAlignment = xlRight
                       
                        
                    End If
                    
                    Volume = Volume + ws.Cells(i, 7).Value
                    ws.Range("N" & SummaryTableRow).Value = Volume
        
                    SummaryTableRow = SummaryTableRow + 1
                    OpeningPrice = ws.Cells(i + 1, 3).Value
                    Volume = 0
                    YearlyChange = 0
                    PercentChange = 0
                
                Else
                
                    Volume = Volume + ws.Cells(i, 7).Value
        
                End If
                
            Next i
            
     End If
        
    Next ws
        
   
    
    
    
End Sub



            
            
