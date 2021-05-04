# Stock_Analysis
Explore green energy stock performance by analyzing financial data using VBA.

## Project Overview
This analysis will examine the trends associated with 12 different stocks:
* AY
* CSIQ
* DQ
* ENPH
* FSLR
* HASI
* JKS
* RUN
* SEDG
* SPWR
* TERP
* VSLR

### DAQO Stock Analysis
In order to investigate whether the green energy stock DAQO (ticker “DQ”) is a sound investment, the stock performance was evaluated based on total daily volume (number of shares traded throughout the day) and yearly return (percent difference in price from the beginning of the year to the end of the year) for years 2017 and 2018.  In order to run analysis for DAQO stock for a given year, a Visual Basic script was created and a button was added to the Excel worksheet “DQ Analysis.”  To run the script, click the button “Run Analysis for DAQO.”  To clear the sheet for a new analysis, click the button “Clear Worksheet.”

### All Stocks Analysis
A larger analysis was performed over all 12 stocks, where performance is based on total daily volume and yearly return for years 2017 and 2018.  Values for each stock are accumulated in the Excel worksheet “All Stocks Analysis”. In order to run analysis for all 12 stocks for a given year, a Visual Basic script was created and a button was added to the Excel Sheet “All Stocks Analysis.”  To run the script, click the button “Run optimized Analysis for All Stocks.”  To clear the sheet for a new analysis, click the button “Clear Worksheet.”

Positive returns are highlighted in green and negative returns are highlighted in red.

Elapsed time for each analysis is calculated and displayed in a pop-up window once a script has been executed.

## Results
This analysis determined that the DAQO stock decreased in value by about 63% in 2018.  For that reason, multiple stocks were examined through All Stocks Analysis to determine the better valued stock.

Running the analysis for all stocks in 2017, only one stock “TERP” decreased in value while all other stocks showed positive returns.  The DQ stock showed positive returns in 2017 and had the lowest Total Daily Volume of shares traded.  Using the “Run optimized Analysis for All Stocks” script, generates the following table:

![VBA_Challenge_2017](/Resources/VBA_Challenge_2017.png "All Stocks Analysis 2017")

Using the optimized script, the elapsed time for the script to run the analysis was only 0.15625 seconds which is significantly less than the original script.

The analysis for all stocks in 2018 is represented in the following table where the majority of stocks had decreased returns, including the DQ stock:  

![VBA_Challenge_2018](/Resources/VBA_Challenge_2018.PNG "All Stocks Analysis 2018")

Using the optimized script, the elapsed time was only 0.140625 seconds which is again, significantly less than the original script.

The stocks ENPH and RUN were the only stocks to show positive returns for both 2017 and 2018.  It was determined that the DAQO stock is volatile and may be a risky investment.  The stocks ENPH and RUN may offer better returns.  It is recommended to utilize the script created in this project to further evaluate stock returns over a longer term if data is available.

## Summary
1.	Code was refactored for the “All Stocks Analysis” macro to optimize the run time.  The refactored script did improve computation time which is an advantage for larger datasets but the disadvantage of an optimized code is that it can sometimes be more difficult to read, understand and debug later.

2.	Once the refactored code for “All Stocks Analysis” was applied, it was clear that the optimized code was running at roughly one third (1/3) the elapsed time of the original code.  Additionally, we were able to populate and format the table in one script.  However, the refactored script utilizes multiple variables as arrays and consecutive “For” loops to achieve this time savings which makes it more difficult for a human to comprehend when reading the script for the first time.  The original code utilizes nested loops to populate the final table which increases computation time but is easier for a human to read and comprehend the script. 
