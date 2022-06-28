' ===========================================================================================
' Module 2 Challenge - Excel VBA Scripting: The VBA of Wall Street
'
'   This script will run on all worksheets. For each worksheet, this will create the
'   following summary tables:
'     1. Stocks "Yearly Stocks Summary" table
'     2. Summary table for the stocks with greatest % increase, greatest % decrease and
'        greatest total volume
'
'   Assumption: The data being processed were sorted by ticker and date in ascending order
'
'   Author      : Rosie Gianan
'   Date Written: June 24, 2022
' ===========================================================================================

' -------------------------------------------------------------------------------------------
' Define the variables that can be accessible to all sub routines
' -------------------------------------------------------------------------------------------
  Dim this_worksheet As Worksheet             ' Variable for the worksheet object
  Dim is_first_ticker As Boolean              ' Variable to identify the first ticker symbol
    
  Dim ticker_symbol As String                 ' Variable for the ticker symbol
  Dim yearly_price_change As Double           ' Variable for the yearly price change
  Dim yearly_price_change_percent As Double   ' Variable for the yearly percent price change
  Dim total_stock_volume As Variant           ' Variable for the total stock volume
   

' -------------------------------------------------------------------------------------------
' VBA_Of_Wall_Street():
'   Main processing routine that will loop through every worksheet. For each worksheet, this
'   will call the sub routine Stocks_Yearly_Summary() passing the worksheet name to process.
'
'   Run this script to generate the two summary tables on the same worksheet as the raw data.
' -------------------------------------------------------------------------------------------
Sub VBA_Of_Wall_Street()

    Dim worksheet_name As String
    
  ' -----------------------------------------------------------------------------------------
  ' Loop through every worksheets. For each worksheet, call the sub routine
  ' Stocks_Yearly_Summary() to create the two summary tables on the same worksheet as the raw
  ' data.
  ' -----------------------------------------------------------------------------------------
    For Each ws_tab In Worksheets

      ' Save the Worksheet Name
        worksheet_name = ws_tab.Name
        
      ' Set worksheet object
        Set this_worksheet = Worksheets(worksheet_name)
        
      ' Create the summary tables for the specific worksheet
        Call Stocks_Yearly_Summary
        
    Next ws_tab    ' Get the next worksheet to process
    
  ' Display processing confirmation message
    MsgBox ("The summary tables were successfully created for all workssheets. Please verify.")

End Sub


' -------------------------------------------------------------------------------------------
' Stocks_Yearly_Summary():
'
' This sub routine will do the the following:
' 1. Loop through all the ticker symbols and creates a summary table with the following data:
'    - Ticker symbol
'    - Yearly price change
'    - Yearly price change percent
'    - Total Stock Volume
'
' 2. Apply conditional formatting to the "Yearly Change ($) and "Percent Change" columns
'    - green for positive change
'    - red for negative
'
' 3. Create a summary table of the ticker symbol(s) with the "greatest" data.
'    - Greatest % Increase
'    - Greatest % Decrease
'    - Greatest Total Volume
' -------------------------------------------------------------------------------------------
 Sub Stocks_Yearly_Summary()

  ' -----------------------------------------------------------------------------------------
  ' Define variables needed for processing
  ' -----------------------------------------------------------------------------------------
    Dim year_first_trading_open_price As Double      ' Variable for the open price of the first trading day of the year
    Dim year_last_trading_closing_price As Double    ' Variable for the closing price of the last trading day of the year
    
    Dim last_row As Long                             ' Variable for the number of rows in the worksheet
    Dim summary_row_num As Long                      ' Variable for the number of rows in the summary table
    Dim range_column As String                       ' Variable to identify which column to apply the conditional formatting
    
  ' -----------------------------------------------------------------------------------------
  ' Initialize the variables
  ' -----------------------------------------------------------------------------------------
  ' Set the row number of the the last record in the worksheet
    last_row = this_worksheet.Cells(Rows.Count, 1).End(xlUp).Row
    
  ' Initialize the yearly summary table row num to 0
    summary_row_num = 0
    
  ' Set first ticker identifier to True
    is_first_ticker = True
  
  ' -----------------------------------------------------------------------------------------
  ' Loop through all the ticker symbols starting at row 2 since row 1 is a header.
  ' For each ticket symbol, save the following information in the yearly summary table:
  '     1. Ticker symbol
  '     2. Yearly price change
  '     3. Yearly price change percent
  '     4. Total Stock Volume
  ' -----------------------------------------------------------------------------------------
    For i = 2 To last_row
    
      ' Save the ticker symbol
        ticker_symbol = this_worksheet.Range("A" & i).Value
    
      ' Create the yearly summary table header once (only on the first ticker)
        If i = 2 Then
        
            this_worksheet.Range("I1").Value = "Ticker"
            this_worksheet.Range("J1").Value = "Yearly Change"
            this_worksheet.Range("K1").Value = "Percent Change"
            this_worksheet.Range("L1").Value = "Total Stock Volume"
            
          ' Add 1 to the number of rows in the summary table. Header is row 1
            summary_row_num = summary_row_num + 1
            
        End If
        
      ' -----------------------------------------------------------------------------------------
      ' Compare the current row vs previous row.
      ' If not the same, then the current row is the first record for the ticker sysmbol.
      ' -----------------------------------------------------------------------------------------
        If ticker_symbol <> this_worksheet.Range("A" & i - 1) Then
            
          ' Save the ticker's open price for the first trading day of the year
            year_first_trading_open_price = this_worksheet.Range("C" & i).Value
          
          ' Initialize the total volume to zero
            total_stock_volume = 0

        End If
        
      ' -----------------------------------------------------------------------------------------
      ' Compare the current row vs next row.
      ' If not the same, then the current row is the last record for the ticker symbol.
      ' If the same, then it is within ticker symbol.
      ' -----------------------------------------------------------------------------------------
        If ticker_symbol <> this_worksheet.Range("A" & i + 1) Then
        
          ' Save the ticker's closing price for the last trading day of the year
            year_last_trading_closing_price = this_worksheet.Range("F" & i).Value
            
          ' Accumulate the total volume count with the current day count
            total_stock_volume = total_stock_volume + this_worksheet.Range("G" & i).Value
            
          ' Calculate the yearly price change
            yearly_price_change = year_last_trading_closing_price - year_first_trading_open_price
       
          ' Calculate the yearly price change percent
            yearly_price_change_percent = yearly_price_change / year_first_trading_open_price
            
          ' Add 1 to the number of rows in the summary table. Add 1 row for each unique ticker
            summary_row_num = summary_row_num + 1
            
          ' Save the ticker's summary information in the worksheet cell location
            this_worksheet.Range("I" & summary_row_num).Value = ticker_symbol
            this_worksheet.Range("J" & summary_row_num).Value = yearly_price_change
            this_worksheet.Range("K" & summary_row_num).Value = yearly_price_change_percent
            this_worksheet.Range("L" & summary_row_num).Value = total_stock_volume
    
          ' -----------------------------------------------------------------------------------------
          ' Save the stocks data in the summary table of the ticker symbol(s) with "greatest" data.
          ' -----------------------------------------------------------------------------------------
            Call Save_Stocks_Greatest_Data
            
       Else
       
          ' Within ticker sysmbol, accumulate the total volume count with the current day count
            total_stock_volume = total_stock_volume + this_worksheet.Range("G" & i).Value
           
       End If

    Next i  'Process the next ticker symbol
    
  ' -----------------------------------------------------------------------------------------
  ' Apply Conditional Formatting to the "Yearly Change" data in Column "J"
  ' -----------------------------------------------------------------------------------------
    range_column = "J"
    Call ConditionalFormatting(range_column, summary_row_num)
    
  ' -----------------------------------------------------------------------------------------
  ' Apply Conditional Formatting to the "Percent Change" data in Column "K"
  ' -----------------------------------------------------------------------------------------
    range_column = "K"
    Call ConditionalFormatting(range_column, summary_row_num)
    
  ' -----------------------------------------------------------------------------------------
  ' Format "Yearly Change" data in column "J" to 2 decimal places
  ' -----------------------------------------------------------------------------------------
    this_worksheet.Range("J:J").NumberFormat = "0.00"
    
  ' -----------------------------------------------------------------------------------------
  ' Format the following cells to percent with 2 decimal places
  ' - "Percent Change" in Column "K"
  ' - Greatest % increase value
  ' - Greatest % decrease value
  ' -----------------------------------------------------------------------------------------
    this_worksheet.Range("K:K").NumberFormat = "0.00%"
    this_worksheet.Range("Q2:Q3").NumberFormat = "0.00%"
    
  ' -----------------------------------------------------------------------------------------
  ' Apply autofit to all column the yearly summary table and the "greatest summary table"
  ' -----------------------------------------------------------------------------------------
    this_worksheet.Columns("I:M").AutoFit
    this_worksheet.Columns("O:Q").AutoFit
    
End Sub


' -----------------------------------------------------------------------------------------
' Save_Stocks_Greatest_Data()
'
' This sub routine will save the stocks data in the summary table of the ticker symbol(s)
' with "greatest" data.
'    - Greatest % Increase
'    - Greatest % Decrease
'    - Greatest Total Volume
' -----------------------------------------------------------------------------------------
Sub Save_Stocks_Greatest_Data()
         
    If is_first_ticker = True Then
                
      ' Set to false so this logic will only run once
        is_first_ticker = False
                                  
      ' Save the "greatest" data header and details description
        this_worksheet.Range("P1").Value = "Ticker"
        this_worksheet.Range("Q1").Value = "Value"
        this_worksheet.Range("O2").Value = "Greatest % Increase"
        this_worksheet.Range("O3").Value = "Greatest % Decrease"
        this_worksheet.Range("O4").Value = "Greatest Total Volume"
                    
      ' The first ticker symbol is assumed to have the "greatest" data. This will be used
      ' as the basis for comparison with the succeeding ticker symbols to find the "actual
      ' greatest" data
        this_worksheet.Range("P2").Value = ticker_symbol                ' Save the ticker symbol with greatest percent increase
        this_worksheet.Range("P3").Value = ticker_symbol                ' Save the ticker symbol with greatest percent decrease
        this_worksheet.Range("P4").Value = ticker_symbol                ' Save the ticker symbol with greatest total Volume
        
        this_worksheet.Range("Q2").Value = yearly_price_change_percent  ' Save the greatest percent increase
        this_worksheet.Range("Q3").Value = yearly_price_change_percent  ' Save the greatest percent decrease
        this_worksheet.Range("Q4").Value = total_stock_volume           ' Save the greatest total volume
        
    Else
      ' Save the greatest percent increase data
        If yearly_price_change_percent > this_worksheet.Range("Q2").Value Then
            this_worksheet.Range("P2").Value = ticker_symbol
            this_worksheet.Range("Q2").Value = yearly_price_change_percent
        End If
        
      ' Save the greatest precent decrease data
        If yearly_price_change_percent < this_worksheet.Range("Q3").Value Then
            this_worksheet.Range("P3").Value = ticker_symbol
            this_worksheet.Range("Q3").Value = yearly_price_change_percent
        End If
        
      ' Save the greatest total stocks volume data
        If total_stock_volume > this_worksheet.Range("Q4").Value Then
            this_worksheet.Range("P4").Value = ticker_symbol
            this_worksheet.Range("Q4").Value = total_stock_volume
        End If
                    
    End If
    
End Sub

' -------------------------------------------------------------------------------------------
' ConditionalFormatting(rangeColumn, lastRow):
'
' Apply the following conditional formatting for a given column and row number:
' - Highlight green for positive
' - Highlight red for negative
'
' This subroutine has the following parameters:
'  1. rangeColumn - Range Column
'  2. lastRow     - last row number
' -------------------------------------------------------------------------------------------
Sub ConditionalFormatting(rangeColumn, lastRow)

    Dim cell_range As Range     'Variable for the cell ranges to apply the formattings

  ' Set the range to format starting at row 2 to exclude the header from formatting
    Set cell_range = this_worksheet.Range(rangeColumn & "2" & ":" & rangeColumn & lastRow)
    
  ' Delete previous conditional formats
    cell_range.FormatConditions.Delete
    
  ' Add the first rule. Highlight green for positive
    cell_range.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreaterEqual, Formula1:="=0"
    cell_range.FormatConditions(1).Interior.Color = vbGreen 'RGB(0, 255, 0)
    
  ' Add the second rule. Highlight red for negative
    cell_range.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="=0"
    cell_range.FormatConditions(2).Interior.Color = vbRed  'RGB(255, 0, 0)

End Sub

' ===========================================================================================
' End of code
' ===========================================================================================










