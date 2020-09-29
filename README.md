# VBAHW
Sub VBAwallStreet()
    For Each ws In Worksheets
    
        Dim LastRow As Double
        Dim startRow As Double
        Dim stockVolume As Double
        Dim ticker As String
        Dim openAmt As Double
        Dim closeAmt As Double
        Dim yearChange As Double
        Dim perChange As Double
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "% Change"
        ws.Range("L1").Value = "Stock Volume"
        

        yearChange = 0
        perChange = 0
        stockVolume = 0
        startRow = 2
    openAmt = ws.Cells(2, 3).Value
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
        For i = 2 To LastRow
            If (ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value) Then
            closeAmt = ws.Cells(i, 6).Value
            openAmt = ws.Cells(i + 1, 3).Value
            yearChange = (closeAmt - openAmt)
            ws.Cells(startRow, 10).Value = yearChange
            perChange = (yearChange / closeAmt)
            ws.Cells(startRow, 11).Value = perChange
            ws.Cells(startRow, 11).NumberFormat = "0.00%"
            ticker = ws.Cells(i, 1).Value
            stockVolume = stockVolume + ws.Cells(i, 7).Value
            ws.Cells(startRow, 9).Value = ticker
            ws.Cells(startRow, 12).Value = stockVolume
            startRow = startRow + 1
            stockVolume = 0
            yearChange = 0
            perChange = 0
        Else
            stockVolume = stockVolume + ws.Cells(i, 12).Value
           
        End If
      Next i
    
    For i = 2 To LastRow
        If ws.Cells(i, 10).Value > 0 Then
            ws.Cells(i, 10).Interior.ColorIndex = 4
        ElseIf ws.Cells(i, 10).Value < 0 Then
            ws.Cells(i, 10).Interior.ColorIndex = 3
        End If
    Next i
     
    Next ws
    
End Sub
