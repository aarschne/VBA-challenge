Attribute VB_Name = "Module1"
Sub Alphabettester():

'Declare variables
Dim CurrentTicker As String
Dim RowCount As Long
Dim FirstOpen As Double
Dim LastClose As Double
Dim CurrentOpen As Double
Dim CurrentClose As Double
Dim CurrentStock As Variant
Dim YearlyChange As Double
Dim TableRow As Integer
Dim PercentChange As Double
Dim StockVolume As Double
Dim NumUniqueTickers As Integer
Dim GreatestPerIncrease As Double
Dim GreatestPerDecrease As Double
Dim GreatestStockVolume As Double
Dim GreatPerIncrTicker As String
Dim GreatPerDecrTicker As String
Dim GreatStockVolTicker As String
Dim ws As Worksheet
Dim starting_ws As Worksheet

'Set starting ws
Set starting_ws = ActiveSheet

'iterate over all the sheets
For Each ws In ThisWorkbook.Worksheets
    
    'activate current sheet
    ws.Activate
    
    'Make Headers
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
    Cells(1, 16).Value = "Ticker"
    Cells(1, 17).Value = "Value"
    Cells(2, 15).Value = "Greatest % Increase"
    Cells(3, 15).Value = "Greatest % Decrease"
    Cells(4, 15).Value = "Greatest Total Volume"
    ActiveSheet.Columns("O").AutoFit
    
    'Initialize tablerow, StockVolume and firstopen
    TableRow = 2
    StockVolume = Cells(2, 7).Value
    FirstOpen = Cells(2, 3).Value
    
    'Initialize number of unique tickers
    NumUniqueTickers = 0
    
    'Find row count
    RowCount = ActiveSheet.UsedRange.Rows.Count
    
    
    'iterate through the filled rows
    For i = 2 To RowCount
    
        'set current variables
        CurrentTicker = Cells(i, 1).Value
        CurrentOpen = Cells(i, 3).Value
        CurrentClose = Cells(i, 6).Value
        CurrentStock = Cells(i, 7).Value
        
        'check to see if the ticker is different
        If CurrentTicker <> Cells(i + 1, 1) Then
        
            'Find the last close
            LastClose = CurrentClose
        
            'Find yearly change, output it and ticker to table
            YearlyChange = LastClose - FirstOpen
            Cells(TableRow, 9).Value = CurrentTicker
            Cells(TableRow, 10).Value = YearlyChange
            
            'Find percentage change, output to summary table
            PercentChange = (LastClose - FirstOpen) / FirstOpen
            Cells(TableRow, 11).Value = Format(PercentChange, "Percent")
            
            'make positive changes green
            If PercentChange > 0 Then
                Cells(TableRow, 11).Interior.ColorIndex = 4
            End If
            If YearlyChange > 0 Then
                Cells(TableRow, 10).Interior.ColorIndex = 4
            End If
            
            'make negative changes red
            If (PercentChange < 0) Then
                Cells(TableRow, 11).Interior.ColorIndex = 3
            End If
            If YearlyChange < 0 Then
                Cells(TableRow, 10).Interior.ColorIndex = 3
            End If
            
            'OutputStockVolume
            Cells(TableRow, 12).Value = StockVolume
        
            'Increment TableRow and Number of unique tickers
            NumUniqueTickers = NumUniqueTickers + 1
            TableRow = TableRow + 1
            
            'Update first open,Stockvolume
            FirstOpen = Cells(i + 1, 3).Value
            StockVolume = Cells(i + 1, 7).Value
            
        Else
            StockVolume = StockVolume + Cells(i + 1, 7).Value
        End If
    Next i
    
    'Autofit the columns
    ActiveSheet.Columns("I:L").AutoFit
    
    'Initialize the variables for finding the maxes and min
    GreatestPerIncrease = Cells(2, 11).Value
    GreatestPerDecrease = Cells(2, 11).Value
    GreatestStockVolume = Cells(2, 12).Value
    GreatPerIncrTicker = Cells(2, 9).Value
    GreatPerDecrTicker = Cells(2, 9).Value
    GreatStockVolTicker = Cells(2, 9).Value
    
    'Find greatest and least percent change and greatest stock volume
    For i = 3 To (NumUniqueTickers + 1)
        If Cells(i, 11).Value > GreatestPerIncrease Then
            GreatestPerIncrease = Cells(i, 11).Value
            GreatPerIncrTicker = Cells(i, 9).Value
        End If
        If Cells(i, 11).Value < GreatestPerDecrease Then
            GreatestPerDecrease = Cells(i, 11).Value
            GreatPerDecrTicker = Cells(i, 9).Value
        End If
        If Cells(i, 12).Value > GreatestStockVolume Then
            GreatestStockVolume = Cells(i, 12).Value
            GreatStockVolTicker = Cells(i, 9).Value
        End If
    Next i
    
    'Output the percent change and stock volume variables
    Cells(2, 16).Value = GreatPerIncrTicker
    Cells(3, 16).Value = GreatPerDecrTicker
    Cells(4, 16).Value = GreatStockVolTicker
    Cells(2, 17).Value = Format(GreatestPerIncrease, "Percent")
    Cells(3, 17).Value = Format(GreatestPerDecrease, "Percent")
    Cells(4, 17).Value = GreatestStockVolume
    
    
    'Format the output of yearly changes to two decimal places
    Range("J2:J" & (NumUniqueTickers + 1)).NumberFormat = "0.00"
    
    Next ws
    
    'reactivate starting sheet
    starting_ws.Activate
End Sub

