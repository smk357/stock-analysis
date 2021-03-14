# stock-analysis
**A VBA Analysis of Green Energy Stock Performance**

## Project Overview

The client is a recent finance graduate seeking to assist his parents in assessing the performance of stocks for 12 green energy firms over 2017 and 2018. The goal of the project was to refactor VBA code that processes daily performance data (daily prices and trade volume) for 12 stocks (sorted by tickers) to compute total daily volume and yearly return on closing price. Code was edited to increase compuational efficiency as indicateded by run time and accuracy of the results.

## Results and Analysis

### Original Code

The dataset consisted of daily prices (open, high, low, closing) and volume for 12 stock tickers for each day of the year the stocks were traded, for two years (2017 and 2018). The code was designed to ouput total daily volume and percentage change in closing price for each ticker between the opening trading day and the final trading day of a particular year.

First, a new worksheet was created to tabulate results. Variables storing the start time and end time were created to measure run time. An input box was created to allow the same code to run the code for any particular year. Sheet headings and header rows (Ticker, Total Daily Volume, Return) were created within VBA using the *range()* and *cells()* functions (e.g. *Range("A1").Value = "All Stocks (" + yearValue + ")"* and *Cells(3, 1).Value = "Ticker"*)

A new array holding stock tickers was created, intitialized, and populated as rows using a *for* loop:

*'Add rows for Ticker column*

*For i = 4 To 15*

*Cells(i, 1).Value = tickers(i - 4)*

*Next i*

The total number of data rows was then determined (*RowCount = Cells(Rows.Count, "A").End(xlUp).Row*). A variable was created and initialized to store the total daily volume. Analysis was subsequently carried out using a main *for* loop running through the tickers and a nested *for* loop running through each data row. *If* conditionals were added so that only data pertaining to that particular ticker was processed. For example, total volume was calculated within the *If* conditional as follows:

*If Cells(j, 1).Value = ticker Then*
    
*totalVolume = totalVolume + Cells(j, 8).Value*
    
*End If*

To determine the starting price and ending closing price for the year, an *If* conditional was used to check if the previous or subsequent row in the tickers column matched the current row, respectively. If the ticker in the previous or subsequent row did not match that of the current row, the current row was selected as the row containing the correct closing price. For example, the starting price was determined as follows:

*If Cells(j, 1).Value = ticker And Cells(j - 1, 1).Value <> ticker Then*
    
*startingPrice = Cells(j, 6).Value*
    
*End If*

The logic follows from the fact that the rows were arranged as blocks for each ticker, with the days of the year going from top to bottom.



