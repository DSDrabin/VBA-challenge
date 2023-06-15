# VBA-challenge
Module 2 Challenge

Sub stock_analysis()
    
    ' Declare and set worksheet
    Dim ws As Worksheet
    
    ' Loop through all worksheets
    For Each ws In Worksheets
    
        ' Create the column headings
        ws.Range("I1").Value = "Ticker"                  ' Ticker symbol column
        ws.Range("J1").Value = "Yearly Change (%)"       ' Yearly change column
        ws.Range("K1").Value = "Percent Change (%)"      ' Percent change column
        ws.Range("L1").Value = "Total Stock Volume"      ' Total stock volume column
    
        ws.Range("P1").Value = "Ticker"                  ' Summary analysis column - Ticker
        ws.Range("Q1").Value = "Value"                   ' Summary analysis column - Value
        ws.Range("O2").Value = "Greatest % Increase"     ' Print - Greatest % Increase label
        ws.Range("O3").Value = "Greatest % Decrease"     ' Print - Greatest % Decrease label
        ws.Range("O4").Value = "Greatest Total Volume"   ' Print - Greatest Total Volume label
    
        ' Define Ticker variable
        Dim Ticker As String
        Ticker = " "
        Dim Ticker_volume As Double
        Ticker_volume = 0
    
        ' Set initial and last row for worksheet
        Dim i As Long
        Dim Lastrow As Long
    
        ' Define Lastrow of worksheet
        Lastrow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
        ' Set new variables for prices and percent changes
        Dim open_price As Double
        open_price = 0
        Dim close_price As Double
        close_price = 0
        Dim price_change As Double
        price_change = 0
        Dim price_change_percent As Double
        price_change_percent = 0
        
        ' Variables for summary analysis
        Dim summary_row As Long
        summary_row = 2 ' Start at row 2 for summary analysis
        
        Dim max_increase As Double
        max_increase = 0
        
        Dim max_decrease As Double
        max_decrease = 0
        
        Dim max_volume As Double
        max_volume = 0
        
        Dim max_increase_ticker As String
        Dim max_decrease_ticker As String
        Dim max_volume_ticker As String
        
        ' Do loop of current worksheet to Lastrow
        For i = 2 To Lastrow
    
            ' Ticker symbol output
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                Ticker = ws.Cells(i, 1).Value
                
                ' Calculate change in Price
                close_price = ws.Cells(i, 6).Value
                price_change = close_price - open_price
                
                ' Populate Ticker symbol in summary analysis
                ws.Range("I" & summary_row).Value = Ticker
                
                ' Populate Yearly Change in summary analysis
                ws.Range("J" & summary_row).Value = price_change / 100
                
                ' Populate Percent Change in summary analysis
                If open_price <> 0 Then
                    price_change_percent = (price_change / open_price) '* 100
                    ws.Range("K" & summary_row).Value = price_change_percent
                Else
                    ws.Range("K" & summary_row).Value = "N/A"
                End If
                
                ' Populate Total Stock Volume in summary analysis
                ws.Range("L" & summary_row).Value = Ticker_volume
                
                ' Check for greatest % increase
                If price_change_percent > max_increase Then
                    max_increase = price_change_percent
                    max_increase_ticker = Ticker
                End If
                
                ' Check for greatest % decrease
                If price_change_percent < max_decrease Then
                    max_decrease = price_change_percent
                    max_decrease_ticker = Ticker
                End If
                
                ' Check for greatest total volume
                If Ticker_volume > max_volume Then
                    max_volume = Ticker_volume
                    max_volume_ticker = Ticker
                End If
                
                ' Reset variables for the next ticker symbol
                Ticker_volume = 0
                open_price = 0
                summary_row = summary_row + 1
                
            Else
                ' Accumulate total stock volume
                Ticker_volume = Ticker_volume + ws.Cells(i, 7).Value
                
                ' Set open price if it hasn't been set yet
                If open_price = 0 Then
                    open_price = ws.Cells(i, 3).Value
                End If
            End If
    
        Next i
        
        ' Populate "Greatest % Increase" result
        ws.Range("P2").Value = max_increase_ticker
        ws.Range("Q2").Value = max_increase
        
        ' Populate "Greatest % Decrease" result
        ws.Range("P3").Value = max_decrease_ticker
        ws.Range("Q3").Value = max_decrease
        
        ' Populate "Greatest Total Volume" result
        ws.Range("P4").Value = max_volume_ticker
        ws.Range("Q4").Value = max_volume
        
        ' Apply conditional formatting to "Yearly Change" column
        Dim yearly_change_formatting As Range
        Set yearly_change_formatting = ws.Range("J2:J" & summary_row - 1)
        
        With yearly_change_formatting.FormatConditions
            ' Formatting for positive change (green)
            .Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="0"
            .Item(.Count).SetFirstPriority
            .Item(1).Interior.Color = RGB(0, 255, 0)
        
            ' Formatting for negative change (red)
            .Add Type:=xlCellValue, Operator:=xlLess, Formula1:="0"
            .Item(.Count).SetFirstPriority
            .Item(1).Interior.Color = RGB(255, 0, 0)
        End With
        
        ' Format the "Yearly Change" column as percentages
        ws.Range("J2:J" & summary_row - 1).NumberFormat = "0.00%"
        
        ' Format the "Percent Change" column as percentages
        ws.Range("K2:K" & summary_row - 1).NumberFormat = "0.00%"
        
        ' Format the "Total Stock Volume" column as number
        ws.Range("L2:L" & summary_row - 1).NumberFormat = "0.00"
        
        ' Format the "Value" column as number
        ws.Range("Q2:Q2" & summary_row - 1).NumberFormat = "0.00%"
        
        ' Format the "Value" column as number
        ws.Range("Q4").NumberFormat = "0.00"
        
        'Apply autowidhth Columns A to Q
        Dim col As Range
            For Each col In ws.Range("A:Q").Columns
                col.EntireColumn.AutoFit
            Next col
        
    Next ws
         
End Sub
