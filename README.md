# stock-analysis
Performing analysis on green energy companies to help Steve advise his parents' portfolio

**Overview of the Project**

The purpose of this analysis is to determine the most appropriate investments for Steve's parents' portfolio based on volume and yearly return. Based on the information provided, the possible components of the portfolio will be limited to 12 green energy stocks.

**Results**

Based on the formulas provided to calculate yearly returns, the 12 stocks included in the analysis performed extremely poorly in 2018 compared to 2017. From a volume perspective, DQ and ENPH had the largest increases in volume YoY. AY and CSIQ had the largest decreases in volume YoY.

The refactored script to dynamically determine the number of stocks in the analysis took approximately 10x as long to run. The difference in execution time was 1 second versus 10 seconds. The screenshots below show the difference in execution times and similarity of results for the refactored script:

![VBA_Challenge_2017](https://user-images.githubusercontent.com/82549092/116828754-b072d500-ab6e-11eb-8742-1bfcccf16fb0.PNG)
![VBA_Challenge_2017-All_Stocks_Refactor](https://user-images.githubusercontent.com/82549092/116828756-b23c9880-ab6e-11eb-9d2c-fe79c12380d8.PNG)

![VBA_Challenge_2018](https://user-images.githubusercontent.com/82549092/116828758-b49ef280-ab6e-11eb-922e-0f6fd3304b66.PNG)
![VBA_Challenge_2018-All Stocks Refactor](https://user-images.githubusercontent.com/82549092/116828759-b5d01f80-ab6e-11eb-9488-da681494b3f0.PNG)


**Summary**
_Advantages and Disadvantages of Refactoring Code in General_




_Advantages and Disadvantages of Refactoring Code in Current VBA Script_
Although the script took longer to run, the following benefits were achieved:
  1. The refactored script does not require the user to enter the universe of equities to include. As the universe grows to include the entire market, adding each equity individually will be time consuming and require constant updates.
  2. 
