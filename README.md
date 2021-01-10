# Stock Analysis

## Overview of Project

The purpose of this project was to expand the dataset we have been working with for the client, Steve. We refactored code in order to apply it efficiently and broadly to many stocks in multiple years. We focused on 12 stocks in particular, and extracted the total daily volume and return of each. 

## Results

A link to the Excel sheet containing the relevant code:
https://github.com/Mikeblanchard/stock-analysis/blob/main/green_stocks.xlsm 

We have included the section of refactored code below:

    '1a) Create a ticker Index
    
    Dim tickerIndex As Integer
    
    '1b) Create three output arrays
    
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    '2a) Create a for loop to initialize the tickerVolumes to zero.
    
    For i = 0 To 11
        
        tickerVolumes(i) = 0
        tickerStartingPrices(i) = 0
        tickerEndingPrices(i) = 0
        
    Next i
        
    '2b) Loop over all the rows in the spreadsheet.
    
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
        
        If Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
        tickerStartingPrices(tickerIndex) = Cells(i, 6).Value

        'End If
        
        End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        
        If Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
        tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
      
        'If  Then
        
        End If
        
        '3d Increase the tickerIndex.
            
        If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
        tickerIndex = tickerIndex + 1
        
        'End If
            
        End If
    
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
    Next i
    
We created arrays and loops, and with the use of If statements, determined the preformance of the 12 stocks in question in 2017 and 2018

2017                                                                    |2018
:----------------------------------------------------------------------:|:------------------------------------------------------------------------------:
![](https://github.com/Mikeblanchard/stock-analysis/blob/main/Resources/Screenshot%202021-01-10%20154016.png) | ![](https://github.com/Mikeblanchard/stock-analysis/blob/main/Resources/Screenshot%202021-01-10%20154932.png)

Drastically different years for the stocks we looked at. 2017 saw all but TERP growing in value, with DQ being the big winner at almost 200% return. 2018 saw all but RUN and ENPH with a postive return. Interesting to note that RUN and ENPH were the only 2 stock that gave a positive return over both years. It would be fascinating to look further back to see how long they've had a proven record of growth. 

## Summary

####    Advantages and Disadvantages 


