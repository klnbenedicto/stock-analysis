## Overview of Project

### Purpose
Steve's parents want to invest in green energy and have already purchased stock in DAQO New Energy Corp (DQ). Steve wants to diversify their funds across other green energy stock as well so has provided raw data on stock performance of 12 green energy stocks, including DQ. This analysis is to look at the history of all 12 green energy stocks over 2017 and 2018 to help inform Steve and his parents on their investment.

## Results
To run analysis on the data, Steve will need to click the "Run Analysis for All" button and the output will show results for each of the 12 energy stocks in question. The first column provides the stock Ticker. To show Steve and his parents how actively each stock was traded, Total Daily Volume provides the sum of the daily volume for each stock for that year. Finally, Return shows the calculated yearly return of each stock. This column represents the increase or decrease in price from the beginning of the year to the end of the year. If the cell is red, it indicates that the investment shrunk and if the column is green, it indicates that the investment grew.

![All Stocks 2017](/Resources/All_Stocks_2017.png)

In 2017, nearly all of the green energy stocks showed a positive return, with the exception if TERP. The stock that grew the most was DQ at 199.4%. If Steve's parents had invested in 2017, they would have made a lot off of their investment.

![All Stocks 2018](/Resources/All_Stocks_2018.png)

The outlook for 2018 was not so rosy. Only two of the green energy stocks yielded a positive return: ENPH and RUN. DQ performed the worst of the 12 with a return of -62.6%.

Refactoring the script resulted in a much quicker run time for both years.

![Original Script 2017 Run Time](/Resources/Module_No_Refactor_2017.png)
![Refactored Script 2017 Run Time](/Resources/VBA_Challenge_2017.png)


![Original Script 2018 Run Time](/Resources/Module_No_Refactor_2018.png)
![Refactored Script 2018 Run Time](/Resources/VBA_Challenge_2018.png)

## Summary

The advantage of refactoring is a quicker run time as a result of more efficient code. By refactoring, the script ran about five times faster. Refactoring also gives an opportunity to eliminate redundancies by looking for ways to not repeat the same commands. 

A disadvantage of refactoring is writing a more complex script that could take longer to debug if you don't know what you're doing. But the more practice, the better one gets.