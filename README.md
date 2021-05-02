# stock-analysis
Performing analysis on green energy companies to help Steve advise his parents' portfolio.

**Overview of the Project**

The purpose of this analysis is to determine the most appropriate investments for Steve's parents' portfolio based on volume and yearly return. Based on the information provided, the possible components of the portfolio will be limited to 12 green energy stocks.

**Results**

Based on the formulas provided to calculate yearly returns, the 12 stocks included in the analysis performed extremely poorly in 2018 compared to 2017. From a volume perspective, DQ and ENPH had the largest increases in volume YoY. AY and CSIQ had the largest decreases in volume YoY.

The refactored script to dynamically determine the number of stocks in the analysis took approximately 9x as long to run. The difference in execution time was aproximately 1 second versus aproximately 9 seconds. The screenshots below show the difference in execution times and similarity of results for the refactored script:

![VBA_Challenge_2017](https://user-images.githubusercontent.com/82549092/116828754-b072d500-ab6e-11eb-8742-1bfcccf16fb0.PNG)
![VBA_Challenge_2017-All_Stocks_Refactor](https://user-images.githubusercontent.com/82549092/116828756-b23c9880-ab6e-11eb-9d2c-fe79c12380d8.PNG)

![VBA_Challenge_2018](https://user-images.githubusercontent.com/82549092/116828758-b49ef280-ab6e-11eb-922e-0f6fd3304b66.PNG)
![VBA_Challenge_2018-All Stocks Refactor](https://user-images.githubusercontent.com/82549092/116828759-b5d01f80-ab6e-11eb-9488-da681494b3f0.PNG)

As a note - The code to loop through the 12 entities is titled _AllStocksAnalysis12StocksRefactored()_ and the refactored code to loop through an inderminate amount of entities is titled _AllStocksAnalysisLoopThroughStocksRefactored_.


**Summary**

_Advantages and Disadvantages of Refactoring Code in General_

In general, refactoring code can have the following advantages and disadvantages in general:
  1. _Advantage:_ Code can be made more efficient through taking fewer steps, using less memory, or improving the logic in general. The original code will be used as a baseline to compare the runtime and results of the updated code.
  2. _Disadvantage:_ Refactoring code can lead to extended runtimes for scripts. This is usually the case when original scripts are not dynamic and have static requirements. Explaining the cause of the increase in runtime to stakeholders is important to ensure the stackholders understand the trade-off between longer runtime and initial resource investment versus shorter runtime and increased long-time script management.
  3. _Advantage:_ Refactoring code also lead to ongoing code reviews that can identify logical errors in the original script. Even though scripts could work when originally developed, requirements and data change over time.

_Advantages and Disadvantages of Refactoring Code in Current VBA Script_

Although a disadvantage of the new script is that it took longer to run, the following benefits were achieved:
  1. The refactored script does not require the user to enter the universe of equities to include. As the universe grows to include the entire market, adding each equity individually will be time consuming and require constant updates.
  2. The refactored script updates the summary page with the exact number of rows in the input year's analysis. The original code required the user to delete rows in the summary page if the new year's population was smaller and also update the code with the new year's tickerIndex + 4 if the population changed. 
  3. (NOT INCLUDED IN CURRENT REFACTOR TO ENSURE VALUES MATCHED FOR CREDIT) - The refactored script should have looked at the logic used to calculate return during the year. Steve is using closing price from the first day of the year instead of closing price from the last day of the previous year. Using closing price from the first day of the year leads to a non-continuous dataset that ignores return on the first day, and could exclude after-hours and pre-market trading during the last day of the previous year and first day of the current year. In addition, adjusted closing price should be used for this analysis, as the adjusted closing price takes into account activity such as stock splits and dividends. If these events occured YoY, his analysis could be based on incorrect data and equities with stock splits or large dividends could be seen as performing poorly.
