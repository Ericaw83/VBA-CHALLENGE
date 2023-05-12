Attribute VB_Name = "Module2"
Sub Ticker()

        'Set variable for worksheets
        Dim ws As Worksheet
        'Loop through all Worksheets
        For Each ws In Worksheets
        'Create the column headings
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"

        'Set an initial variables
        Dim Ticker As String
        
        Dim SummaryTableRow As Integer
        
        SummaryTableRow = 2
        
        Dim TotalVolume As Double
        
        TotalVolume = 0
        
        Dim OpenPriceRow As Integer
        
        OpenPriceRow = 2
        
        'Define Lastrow of worksheet
        lastLine = ws.Cells(Rows.Count, "A").End(xlUp).Row
        
        'Loop through to Lastrow
        For i = 2 To lastLine
        
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
        
        Ticker = ws.Cells(i, 1).Value
        
        ws.Range("I" & SummaryTableRow).Value = Ticker
        
        
        OpenPrice = ws.Cells(OpenPriceRow, 3).Value
        
        ClosePrice = ws.Cells(i, 6).Value
        
        YearlyChange = ClosePrice - OpenPrice
        
        PercentChange = YearlyChange / OpenPrice
        
        ws.Range("J" & SummaryTableRow).Value = YearlyChange
        
        ws.Range("J" & SummaryTableRow).NumberFormat = "0.00"

        
        ws.Range("K" & SummaryTableRow).Value = PercentChange
        
        ws.Range("K" & SummaryTableRow).NumberFormat = "0.00%"

        
        
        TotalVolume = TotalVolume + ws.Cells(i, 7).Value
        
        ws.Range("L" & SummaryTableRow).Value = TotalVolume
        
        If YearlyChange > 0 Then
        
        ws.Range("J" & SummaryTableRow).Interior.ColorIndex = 4
        
        ElseIf YearlyChange < 0 Then
        
        ws.Range("J" & SummaryTableRow).Interior.ColorIndex = 3
        
    
        End If
        
        SummaryTableRow = SummaryTableRow + 1
        
        TotalVolume = 0
        
        YearlyChange = 0
        
        Else
        
        TotalVolume = TotalVolume + ws.Cells(i, 7).Value
    
        
        End If
                
   Next i
   
   
   Next ws
    
End Sub
