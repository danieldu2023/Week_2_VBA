# Week_2_VBA

'Code inspired from https://github.com/emmanuelmartinezs/stock-analysis

Sub Multi_Year_Stock_Data()
    
    'define variable type'
    Dim ws As Worksheet
    Dim select_index As Double
    Dim first_row As Double
    Dim select_row As Double
    Dim last_row As Double
    Dim year_opening As Single
    Dim year_closing As Single
    Dim volume As Double

    'loop through each work sheet
    For Each ws In Sheets
        Worksheets(ws.Name).Activate
        select_index = 2
        first_row = 2
        select_row = 2
        last_row = WorksheetFunction.CountA(ActiveSheet.Columns(1))
        volume = 0
        
     'assign headers to columns
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Yearly Change"
        Cells(1, 11).Value = "Percent Change"
        Cells(1, 12).Value = "Total Volume"
        Cells(1, 16).Value = "Ticker"
        Cells(1, 17).Value = "Value"
        
        Cells(2, 15).Value = "Greatest % Increase"
        Cells(3, 15).Value = "Greatest % Decrease"
        Cells(4, 15).Value = "Greatest Total Volume"
        
        
      'find all unique tickers
        For i = first_row To last_row
            tickers = Cells(i, 1).Value
            tickers2 = Cells(i - 1, 1).Value
            If tickers <> tickers2 Then
                Cells(select_row, 9).Value = tickers
                select_row = select_row + 1
            End If
         Next i
    
       'calculating total volumes
        For i = first_row To last_row + 1
            tickers = Cells(i, 1).Value
            tickers2 = Cells(i - 1, 1).Value
            If tickers = tickers2 And i > 2 Then
                volume = volume + Cells(i, 7).Value
            ElseIf i > 2 Then
                Cells(select_index, 12).Value = volume
                select_index = select_index + 1
                volume = 0
            Else
                volume = volume + Cells(i, 7).Value
            End If
        Next i
            
       'finding year open price and year closing price
        select_index = 2
        For i = first_row To last_row
            If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
                year_closing = Cells(i, 6).Value
            ElseIf Cells(i, 1).Value <> Cells(i - 1, 1).Value Then
                year_opening = Cells(i, 3).Value
            End If
        'finding yearly change
            If year_opening > 0 And year_closing > 0 Then
                increase = year_closing - year_opening
                percent_increase = increase / year_opening
                Cells(select_index, 10).Value = increase
        'formating yearly change to percent
                Cells(select_index, 11).Value = FormatPercent(percent_increase)
                year_closing = 0
                year_opening = 0
                select_index = select_index + 1
            End If
        Next i
        
        'define max and min for greatest increase and decrease
        max_per = WorksheetFunction.Max(ActiveSheet.Columns("k"))
        min_per = WorksheetFunction.Min(ActiveSheet.Columns("k"))
        max_vol = WorksheetFunction.Max(ActiveSheet.Columns("l"))
        'formating greatest increase and decrease to percent
        Range("Q2").Value = FormatPercent(max_per)
        Range("Q3").Value = FormatPercent(min_per)
        Range("Q4").Value = max_vol
        
        
        'finding greatest increase and decrease
        For i = first_row To last_row
            If max_per = Cells(i, 11).Value Then
                Range("P2").Value = Cells(i, 9).Value
            ElseIf min_per = Cells(i, 11).Value Then
                Range("P3").Value = Cells(i, 9).Value
            ElseIf max_vol = Cells(i, 12).Value Then
                Range("P4").Value = Cells(i, 9).Value
            End If
        Next i
        
        'formating yearly change colum according the right color
        For i = first_row To last_row
            If IsEmpty(Cells(i, 10).Value) Then Exit For
            If Cells(i, 10).Value > 0 Then
                Cells(i, 10).Interior.ColorIndex = 4
            Else
                Cells(i, 10).Interior.ColorIndex = 3
            End If
        Next i
    Next ws
                
End Sub
 'Code inspired from https://github.com/emmanuelmartinezs/stock-analysis
 
