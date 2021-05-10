# stock-analysis
Performing analysis on stock data

## Overview of Project
### Background
The stock data consisted of stock information on 12 clean energy stocks from the year 2017 and 2018. The stock information included the ticker value,    date stock was issued, opening price, highest price, lowest plrice, closing price, adjusted closing price, and stock volume. 

### Purpose
The purpose of this analysis was to refactor a Microsoft Excel VBA code in order to measure the performance of clean energy stocks over the last few years and to determine whether refactoring my code successfully made the VBA script run faster.

## Results
### Stock Performance
I wrote a VBA script by using the green_stocks dataset to loop through the data over the last few years to output the “Ticker,” “Total Daily Volume,” and “Return” columns in your spreadsheet.

#### 2017 Performance
![all_stocks_analysis_2017_refactored.png](https://github.com/yessiez/stock-analysis/blob/master/Resources/all_stocks_analysis_2017_refactored.png?raw=true)

#### 2018 Performance
![all_stocks_analysis_2018_refactored.png](https://github.com/yessiez/stock-analysis/blob/master/Resources/all_stocks_analysis_2018_refactored.png?raw=true)

### Execution Times
I wrote a VBA script to calculate how long the code takes to execute and output the elapsed time in a message box for each year.

#### 2017 Execution Times
##### Original
![year_value_analysis_2017.png](https://github.com/yessiez/stock-analysis/blob/master/Resources/year_value_analysis_2017.png?raw=true)

##### Refactored
![VBA_Challenge_2017.png](https://github.com/yessiez/stock-analysis/blob/master/Resources/VBA_Challenge_2017.png?raw=true)


#### 2018 Execution Times
##### Original
![year_value_analysis_2018.png](https://github.com/yessiez/stock-analysis/blob/master/Resources/year_value_analysis_2018.png?raw=true)

##### Refactored
![VBA_Challenge_2018.png](https://github.com/yessiez/stock-analysis/blob/master/Resources/VBA_Challenge_2018.png?raw=true)

Based on the run-times, the refactored code is more efficient.

### Refactored Script
```
    tickerIndex = tickers(x)
    
    Dim tickerVolumes As Long
    Dim tickerStartingPrices As Single
    Dim tickerEndingPrices As Single
    
    Dim i
    For i = 0 To 11
        tickerVolumes = 0
    
    'Activate data worksheet
    Worksheets(yearValue).Activate
    
    Dim j
    For j = 2 To RowCount
   
        If Cells(j, 1).Value = tickerIndex Then
            tickerVolumes = tickerVolumes + Cells(j, 8).Value
        End If
        
        If Cells(j - 1, 1).Value <> tickerIndex And Cells(j, 1).Value = tickerIndex Then
            tickerStartingPrices = Cells(j, 6).Value
        End If
       
        If Cells(j + 1, 1).Value <> tickerIndex And Cells(j, 1).Value = tickerIndex Then
            tickerEndingPrices = Cells(j, 6).Value
        End If
        
        Next j
    
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = tickerIndex
        Cells(4 + i, 2).Value = tickerVolumes
        Cells(4 + i, 3).Value = tickerEndingPrices / tickerStartingPrices - 1
           
            tickerIndex = tickers(i + 1)

```

## Summary

Refactoring code improves helps us debug programs and also helps in executing the program more efficiently. It also helps you approach a problem differently. A disadvantage is that you can eaisily get lost. In regard to this VBA script, I had a lot of trouble with the refactored code and I can't think of an advantage, only disadvantages.  

