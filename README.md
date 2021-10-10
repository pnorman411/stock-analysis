# Stock Analysis

## Overview of the Project

### Purpose

The purpose of this project is to analyze green energy stocks using Visual Basic for Applications (VBA).

### Background

My client is interested in investing all of their money into DAQO (Ticker: DQ) stocks. Upon analyzing the DQ's 2018 data and concluding DQ had a negative return, it was determined that additional green energy stocks should be analyzed to provide alternatives to the client who may want to diversify their portfolio.

### Refactoring Code

The client was so impressed with the ability to calculate the total volume and annual return for the 12 stocks that they would like to use the code to analyze additional stocks.  I refactored the code to determine if the VBA scripts would run faster if thousands of stocks needed to be analyzed.

## Results

DQ had a positive return of 199% in 2017 but plummeted in 2018 for a -62.6% return.  Green energy stocks as a whole had a successful year in 2017 based on returns.  For 2018, ENPH and RUN were the only two green energy stocks that calculated positive returns.

## Summary

### Advantages and Disadvantages of Refactoring Code

Refactoring code is advantageous because it is intended to make the code run more efficiently which takes fewer steps, uses less memory, and improves the logic for future readers.

The disadvantages of refactoring code are that any time you change the code, you are introducing the risk of additional debugging and there is an increase in time commitment and resources needed to refactor the code. 

### Advantages and Disadvantages of the Original VBA and Refactored Scripts

#### The Original VBA Script

The original VBA script is easier to follow for novice coders because it is direct as to which property or variable is being referenced, e.g. startingPrice and totalVolume are clearly defined instead of having nested definitions.

The disadvantage of the original VBA script is that it has nested loops which can be hard to follow and take longer to loop through the data.

#### The Refactored Script

An advantage of the refactored script is that it ran faster than the original script.  The 2017 data ran 83.6% faster (0.1367188 seconds versus 0.8359375 seconds), and the 2018 data ran 84.7% faster (0.125 seconds versus 0.8164063 seconds).  Running thousands of tickers through the refactored scripts could save a significant amount of time.

The refactored script needed much debugging and trial and error in order to get it to run.  It may be hard for some to follow and comprehend the nested definitions, e.g. tickers(tickerIndex), tickerStartingPrices(tickerIndex), etc.