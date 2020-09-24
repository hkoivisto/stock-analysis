# VBA of Wall Street

## Overview

The purpose of this project is to review raw stock trading data and compile it into an accurate and visually consise annual summary. For this project, 12 stock tickers of interest are reviwed: AY, CSIQ, QG, ENPH, FSLR, HASI, JKS, RUN, SEDG, SPWR, TERP, and VSLR.

## Results

### Analysis

The data in this analysis is divided by year, with an individual year's worth of data on each worksheet. For each year, there is information for each trading day, including opening and closing prices, high and low prices, and total volume for that day. The user specifies one year at a time to analyze. The sript cycles thorugh each line of data in the selcted worksheet, and sums the following values for the year: total trading volume, initial trading price, and ending trading price. The Summary output tabulates this value and also calculates the percentage of change of the ending price over the starting price.

An example of the FOr loop to capture the entire worksheet:
'''
For i = 2 To RowCount

    '3a) Increase volume for current ticker
        If Cells(i, 1).Value = tickers(tickerIndex) Then
            tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
            End If
        
    '3b) Check if the current row is the first row with the selected tickerIndex.
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            End If
            
    '3c) check if the current row is the last row with the selected ticker
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            
    '3d Increase the tickerIndex.
            tickerIndex = tickerIndex + 1    
        End If
    Next i
'''

### Findings

With the exception of ticker RUN, all stocks performed more favorably in 2017 than in 2018. This could indicate a downturn in the clean energy market overall for 2018. Based on these data, the best investment potential is in ticker ENPH, which is the only stock to perform favorably both years. 

   ##### 2017 Tabulated Reults
      
![2017_Stock_Results](https://github.com/hkoivisto/stock-analysis/blob/master/Resources/2017_Stock_Results.png)

   ##### 2018 Tabulated Reults
      
![2018_Stock_Results](https://github.com/hkoivisto/stock-analysis/blob/master/Resources/2018_Stock_Results.png)

### Execution Times

This code was written in two stages. The intial programming resulted in the functionality described above, but required the code to loop through every row of the worksheet 12 times, once for each ticker. This resulted in a script that could run the 2017 worksheet in 1.21875 seconds and could run the 2018 worksheet in 1.195313 seconds. This was determined by a built in time function within th e code.

The code was then refactored to increase efficiency. The new method, as shown in the excerpt above, loops thorugh each line of the worksheet only once to make the same evaluation. THis results in code that can runthe 2017 and 2018 worksheets in and .3164063 .3125 seconds, respectively. This is an average increased efficiency of 74%.

Timer results:

![VBA_Challenge_2017](https://github.com/hkoivisto/stock-analysis/blob/master/Resources/VBA_Challenge_2017.png)

![VBA_Challenge_2018](https://github.com/hkoivisto/stock-analysis/blob/master/Resources/2018_Stock_Results.png)


## Summary

1. What are the advantages or disadvantages of refactoring code?
  - Refactoring code can result in more efficiant analysis, while retaining the original functionalityof a script.
  - Increasing efficiency in a set of code allows the end user to run analysis on increasingly larger sets of data without significant impact to the run time.
  - A potential disadvantage to refactor code is the possibility of unintentionally ceating an error in a previous working code. Care should be taken to not change the basic functionality of the original code.

2. How do these pros and cons apply to refactoring the original VBA script?
  - The refactored version of our All Stocks Analysis ran siginificantly faster than the original code. THis would allow us to perform similar anlaysis on additioanl Tickers without adding too much time to the process. The orignal code took 4 times as long to process as the refactored code. This factor could would be multiplied as hundreds or thousands of additioanl Tickers are added for consideration.
  - A drawback of the refactored code in this project is that the data to be analyzed must have tickers in the order origially defined in the tickers array. The original code could pull and label tickers in any order.
