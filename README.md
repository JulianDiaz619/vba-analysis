# VBA Stocks Analysis

## Overview of project
The purpose of this analysis is to refactor a VBA script with the goal of making it more efficient.The VBA script has the job of separating all the tickers and finding the total volume and yearly return of each company for the year.The run time of the original code is around ~0.5 seconds, so we are hoping to reduce that significantly with the refactor.
## Results
#### Ticker Index

I began by declaring a variable "TickerIndex" set equal to zero. This will help us in our loop later on.

#### Arrays
![Array](array.png)

Here are three output arrays for tickerVolumes, tickerStartingPrices, and tickerEndingPrices. tickerVolumes will store each tickers total volume, tickerStartingPrices will get the starting price, and tickerEndingPrices will get the ending price so that we are able to calculate that years return of each company.
#### tickerVolumes to zero

img

Here is the first for loop to initialize our tickerVolumes array so that all objects in the array are equal to zero.
#### Looping over the rows
img

Here is a loop that loops over every row and fills the tickerVolumes array with each ticker having its totalVolume.It also checks whether the row selected is either a starting or ending price, then puts them in their according arrays for every ticker.
#### Formatting
For the formatting, another for loop was used to add tickers to the column in the all stocks analysis sheet, count the total volume for each ticker, and give us the return for the year. Below is some more formatting to give the chart some visual representation. 
Here is how fast the original code work versus the refactored code.

imge

## Summary
An advantage of refactoring code is that you are able to take already existing code and make it more efficient. There is more to gain from proper code refactoring than to not do it. A disadvantage is just the fact that it is time consuming and the code may work fine without it.
