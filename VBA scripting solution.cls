VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub homework()
For Each ws In Worksheets

    'create the summary table format
    
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    
    ws.Range("N2").Value = "Greatest % Increase"
    ws.Range("N3").Value = "Greatest % Decrease"
    ws.Range("N4").Value = "Greatest Total Volume"
    ws.Range("O1").Value = "Ticker"
    ws.Range("P1").Value = "Value"
    
    'change column width
    ws.Columns("J").ColumnWidth = 12
    ws.Columns("K").ColumnWidth = 13
    ws.Columns("L").ColumnWidth = 16
    ws.Columns("N").ColumnWidth = 20
    
    'make total rows
    total_rows = ws.Cells(2, 1).End(xlDown).Row
    'create total stock volume variable
    Dim total_stock_volume As LongLong
    
    
    
    'create year change variable
    Dim year_change As Double
    
    'create loop to determine ticker changes
    summary_row = 2
    
    For rownum = 2 To total_rows
        
        
        If ws.Cells(rownum, 1).Value <> ws.Cells(rownum + 1, 1).Value Then
        'add tickers to summary table
        ws.Cells(summary_row, 9).Value = ws.Cells(rownum, 1).Value
         'intially define opening price
        open_price = Cells(2, 3).Value
        
        
        'add to the total stock volume
        total_stock_volume = total_stock_volume + ws.Cells(rownum, 7).Value
        'add totalstock volume to summary table
        ws.Cells(summary_row, 12).Value = Format(total_stock_volume, "standard")
        
        'define closing price for each ticker
        close_price = ws.Cells(rownum, 6).Value
        'calculate yearly change
        year_change = close_price - open_price
        'add yearly change to summary table
        ws.Cells(summary_row, 10).Value = year_change
        
        'calculate percent change
            If open_price <> 0 Then
                Percent_Change = (year_change / open_price)
            Else
                Percent_Change = (year_change / 1)
            End If
        
        
        'add percent change to summary table
        ws.Cells(summary_row, 11).Value = FormatPercent(Percent_Change)
        
        'add to the summary table row number
        summary_row = summary_row + 1
         
        'Reset total stock volume to 0
        total_stock_volume = 0
        
        'reset opening price for new ticker
        open_price = Cells(rownum + 1, 3).Value
        Else
        'add to the total stock volume
        total_stock_volume = total_stock_volume + ws.Cells(rownum, 7).Value
        End If
        
       
    Next rownum
    
    'conditional formatting of year change
    For rownum = 2 To total_rows
        If ws.Cells(rownum, 10).Value > 0 Then
        'positive year change set to green
        ws.Cells(rownum, 10).Interior.ColorIndex = 4
        
        Else
        'negative year change set to red
        ws.Cells(rownum, 10).Interior.ColorIndex = 3
        
        End If
    
    
    Next rownum
    
'attempt at filling in hard solution
Dim percent_max As Double
Dim percent_min As Double
    percent_max = Application.WorksheetFunction.Max(Columns("K"))
    ws.Cells(2, 15).Value = FormatPercent(percent_max)
    
    percent_min = Application.WorksheetFunction.Min(Columns("K"))
    ws.Cells(3, 15).Value = FormatPercent(percent_max)
Next ws

End Sub
