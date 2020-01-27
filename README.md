# Stock Data Analysis
## Introduction
The Excel Macro Enabled file in this repository contains stock data for the alternative energy companies in the 2017-2018 years.
For each company the following information is provider: **Ticker**, **Date**, **Open** price, **High** price, **Low** price, **Close** 
price, **Adj Close** price and **Volume** of all stocks sold this day.

We wrote macros (subroutines) in VBA calculating the following data for every ticker (company):
* **Total Daily Volume** shows how many stocks were sold through the year
* **Return** showing if the stocks price for the given company decreased or increased by the end og the year anb at what percent

After puting the analysis data into a table we format it to highlight the positive and negative return with different colors 
for better presentation. Every output worksheet contains buttons allowing to run analysis for the data from a chosen worksheet (year) 
and to clear all the cells in the worksheet.

Every Stock Analysis macro does the same analysis and gives the same output however the code is different to show that the same task 
can be solved in many different ways (some of them more efficient, some - less efficient).

## FormatOutputWorksheet() macro
I've put the initial formatting into a separate subroutine _FormatOutputWorksheet(worksheetName As String, year As String)_ that
can be called from another subroutine and takes as arguments worksheet name and a year.
This macro creates data analysis header and formats the header and the outcome sells.

## AllStockAnalysis() macro
_AllStockAnalysis()_ macro does calculations as described in chapter 2.3.3.

After formatting an output worksheet using another macro _FormatOutputWorksheet()_ we create an array containing all the tickers 
and manually initialize it with ticker names.
Then macro calculates number of all the stock data rows using code we've found online (stack-overflow).
Then this macro loops through the stock data rows as many times as we have tickers to collect/calculate the analysis data for
every ticker in a tickers array and put the data in the output worksheet.

After that we perform conditional formatting on the received analysis data to highlight the positive and negative return data 
with different colors for better presentation.

## ChallengeAllStockAnalysis() macro
The _ChallengeAllStockAnalysis()_ subroutine contains the same calculations as _AllStockAnalusis()_ but the code is refactored 
to loop through the stock data rows only once and store the received data (total volume, starting and ending prices and the current 
ticker) in the corresponding arrays.

After running through all the stock data rows we put the analysis data from the arrays into the ouput worksheet and perform 
conditional formatting on it.

## MyAllStockAnalysis() macro
I've written this macro before reading the chapter 2.3 to present my own vision of how to perform the required ananlysis. 
I think it's inefficient to manually put all the companies' tickers into an array: there might be so many rows that it would be 
impossible to manually go through all of those to collect tickers. I suggest to not use tickers array and collect tickers names 
on the run as all the other data.
It is also inefficient to use nested loops to go through all the rows as many times as the number of companies we have
while we can loop through all the rows just once and collect all the neccessary data including companies' tickers and put that data 
to the output worksheet on the run.

In the one loop that goes from top to bottom through all the stock data rows I have 3 conditionals:
* if we meet the ticker in a current row for the first time - I set the starting price value
* for every row I increase Total Volume value which is reseted before the loop and when we reach the last mention for the current ticker
* if we meet the ticker in a current row for the last time - I set the ending price value and also put all the ananlysis data values
(current ticker, total volume and starting and ending price) into the output worksheet and them reset those values.

After running through all the stock data rows I perform the same conditional formatting on the analysis data.

