Attribute VB_Name = "Module1"
Sub stock_market():

' Defining variables for the first challange
Dim ticker As String
Dim ticker_count As Integer
Dim last_row As Long

Dim opening_price As Double
Dim closing_price As Double

Dim yearly_change As Double
Dim percent_change As Double
Dim stock_volume As Double


' Create a Loop that will include all sheets in this workbook
For Each ws In Worksheets

    ' Make the worksheet active.
    ws.Activate
    
  ' Define Headers
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    
    ' Identify the last row in each sheet according to the 1st column
    last_row = ws.Cells(Rows.Count, "A").End(xlUp).Row
    
    ' Set the initial value of each variable
    ticker_count = 0
    ticker = ""
    yearly_change = 0
    opening_price = 0
    percent_change = 0
    total_stock_volume = 0
    
    ' Loop through the list of tickers
    For i = 2 To last_row
    
        ticker = Cells(i, 1).Value
        
        ' Defining opening price of each ticker
        If opening_price = 0 Then
            opening_price = Cells(i, 3).Value
        End If
        
        ' Add up the stock volume value per ticker as we run through the loop
        stock_volume = stock_volume + Cells(i, 7).Value
        
        ' identifies change of ticker in the loop
        If Cells(i + 1, 1).Value <> ticker Then
            ' Increase the ticker count
            ticker_count = ticker_count + 1
            ' Sets each ticker on the new ticker column
            Cells(ticker_count + 1, 9) = ticker
            
            ' Setting yearly_change
            closing_price = Cells(i, 6)
            yearly_change = closing_price - opening_price
            
            ' Add the yearly change value per ticker to the approriate new column
            Cells(ticker_count + 1, 10).Value = yearly_change
            
            ' Setting color changes on yearly change colum
            If yearly_change > 0 Then
                Cells(ticker_count + 1, 10).Interior.ColorIndex = 4
            ElseIf yearly_change < 0 Then
                Cells(ticker_count + 1, 10).Interior.ColorIndex = 3
            Else
                Cells(ticker_count + 1, 10).Interior.ColorIndex = 6
            End If
            
            
            ' Calculate percent change value for ticker.
            If opening_price = 0 Then
                percent_change = 0
            Else
                percent_change = (yearly_change / opening_price)
            End If
            
            
            ' Format the percent_change value as a percent while keeping it as a double
            Cells(ticker_count + 1, 11).Value = percent_change
            Cells(ticker_count + 1, 11).NumberFormat = "0.00%"
         
            
            ' Reset opening price when ticker changes
            opening_price = 0
            
            ' Add the stock volume value per ticker to the approriate new column
            Cells(ticker_count + 1, 12).Value = stock_volume
            
            ' Reset total stock volume when ticker changes
            stock_volume = 0
        End If
        
    Next i
    
    ' Challenge 2
    ' Defining variables for the second part of the challange
    Dim greatest_percent_increase As Double
    Dim greatest_percent_increase_ticker As String
    Dim greatest_percent_decrease As Double
    Dim greatest_percent_decrease_ticker As String
    Dim greatest_stock_volume As Double
    Dim greatest_stock_volume_ticker As String

    'Create the labels for the values we are identifying
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
    
    ' Set the last row variable for our new loop in the new column I
    last_row = ws.Cells(Rows.Count, "I").End(xlUp).Row
    
    
    ' Start each one of the variables with the respective values in the begining of our loop

    greatest_percent_increase = Cells(2, 11).Value
    greatest_percent_increase_ticker = Cells(2, 9).Value
    greatest_percent_decrease = Cells(2, 11).Value
    greatest_percent_decrease_ticker = Cells(2, 9).Value
    greatest_stock_volume = Cells(2, 12).Value
    greatest_stock_volume_ticker = Cells(2, 9).Value
    
    
    
    ' Loop through our new list of tickers
    For i = 2 To last_row
    
        ' Find the ticker with the greatest percent increase.
        If Cells(i, 11).Value > greatest_percent_increase Then
            greatest_percent_increase = Cells(i, 11).Value
            greatest_percent_increase_ticker = Cells(i, 9).Value
        
        ElseIf Cells(i, 11).Value < greatest_percent_decrease Then
            greatest_percent_decrease = Cells(i, 11).Value
            greatest_percent_decrease_ticker = Cells(i, 9).Value
        End If
        
        ' Find the ticker with the greatest stock volume.
        If Cells(i, 12).Value > greatest_stock_volume Then
            greatest_stock_volume = Cells(i, 12).Value
           greatest_stock_volume_ticker = Cells(i, 9).Value
        End If
        
    Next i
    
    ' Add the values identified to the respective cells for display
    Range("Q2").Value = Format(greatest_percent_increase, "Percent")
    Range("P2").Value = Format(greatest_percent_increase_ticker, "Percent")
    Range("P3").Value = Format(greatest_percent_decrease_ticker, "Percent")
    Range("Q3").Value = Format(greatest_percent_decrease, "Percent")
    Range("P4").Value = greatest_stock_volume_ticker
    Range("Q4").Value = greatest_stock_volume
    
 
 
Next ws


End Sub


