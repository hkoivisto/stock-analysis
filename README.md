# VBA of Wall Street

## Overview

THe purpose of this project is to review raw stock trading data and compile it into an accurate and visually consie annual summary. For this project, 12 stock tickers of interest are reviwed: AY, CSIQ, QG, ENPH, FSLR, HASI, JKS, RUN, SEDG, SPWR, TERP, and VSLR.

## Results

### Analysis

The data in this analysis is divided by year, witch an individual year's worth of data on each worksheet. For each year, there is information for each trading day, including opening and closing prices, high and low prices, and total volume for that day. The end user specifies one year at a time to analyze. THe sript cycles thorugh each line of data in the selcted worksheet, and sums total trading volume, initial trading price, and ending trading price for the entire year. The Summary output tabulates this value and also calculates the percentage of change of hte ending price over the starting price.

### Findings

With the exception of ticker RUN, all stocks performed for favorably in 2017 than in 2018. This could indicate a downturn in the clean energy market overall for 2018. Based on these data, the best investment potential is in ticker ENPH, which is the only stock to perform favorably both years. 

      ##### 2017 Tabulated Reults
      
![2017_Stock_Results](https://github.com/hkoivisto/stock-analysis/blob/master/Resources/2017_Stock_Results.png)

      ##### 2018 Tabulated Reults
      
![2018_Stock_Results](https://github.com/hkoivisto/stock-analysis/blob/master/Resources/2018_Stock_Results.png)

## Summary

1. What are the advantages or disadvantages of refactoring code?
  - Refactoring code can result in more efficiant analysis, while retaining the original functionalityof a script.
  - Increasing efficiency in a set of code allows the end user to run analysis on increasingly larger sets of data without significant impact to the run time.
  - A potential disadvantage to refactor code is the possibility of unintentionally ceating an error in a previous working code. Care should be taken to not change the basic functionality of the original code.

2. How do these pros and cons apply to refactoring the original VBA script?
  - The refactored version of our All Stocks Analysis ran siginificantly faster than the original code. THis would allow us to perform similar anlaysis on additioanl Tickers without adding too much time to the process. The orignal code took 4 times as long to process as the refactored code. This factor could would be multiplied as hundreds or thousands of additioanl Tickers are added for consideration.
  - Refactor this code did present challenges to ensue that all variables were defined correctly, because the original varibales and new variables were very similar.
