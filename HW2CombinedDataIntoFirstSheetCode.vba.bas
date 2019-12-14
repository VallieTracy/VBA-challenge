Attribute VB_Name = "CombinedDataIntoFirstSheet"
Sub CombinedData():
    Sheets.Add.Name = "Combined Data"
    Sheets("Combined Data").Move Before:=Sheets(1)
    Set CombinedSheet = Worksheets("Combined Data")
    For Each ws In Worksheets
    LastRow = CombinedSheet.Cells(Rows.Count, "A").End(xlUp).Row + 1
    LastRowTableTicker = ws.Cells(Rows.Count, "K").End(xlUp).Row - 1
    CombinedSheet.Range("A" & LastRow & ":D" & ((LastRowTableTicker - 1) + LastRow)).Value = ws.Range("K2:N" & (LastRowTableTicker + 1)).Value
    Next ws
    CombinedSheet.Range("A1:D1").Value = Sheets(3).Range("K1:N1").Value
    CombinedSheet.Range("A1:D1").Font.Bold = True
    CombinedSheet.Columns("A:D").AutoFit
    CombinedSheet.Range("C:C").NumberFormat = "0.00%"
        
    CombinedDataLastRow = CombinedSheet.Cells(Rows.Count, 1).End(xlUp).Row
    
    For i = 2 To CombinedDataLastRow
        If Cells(i, 2).Value > 0 Then
        Cells(i, 2).Interior.ColorIndex = 4
        ElseIf Cells(i, 2).Value = 0 Then
        Cells(i, 2).Interior.ColorIndex = 33
        ElseIf Cells(i, 2).Value < 0 Then
        Cells(i, 2).Interior.ColorIndex = 3
        
        End If
        
    
    Next i
    

End Sub

