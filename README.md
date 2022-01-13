# Stock analysis for Steve
## *Stock-analysis module 2 assignment*
## Project Overview: Purpose of this analysis
This project creates a report to analyze green stocks performance in 2017 and 2018.  The report is for Steve to advise his parents on investing in worthy "green" stocks.  The report was created with Excel, Visual Basic Application (VBA) and macros.  The initial report focused on stock DQ, a stock his parents were interested in. The final report analyzes an additional 11 "green" stock in terms of volume, starting price and ending price.
## Results
The first report produced results, but seemed to have room for improvement in terms of more elegant programming,
less lines of  codes and more efficient run times.  Although it executed and produced accurate report results, by refactoring the code, run times were reduced signficantly, whilst producing the same accurate results.
Refactoring included creation of arrays to reduce nested loop execution times.   For example, a snippet of the inital code with nested loops looked like this:

'   *Loop through tickers*
   For i = 0 To 11
       ticker = tickers(i)
       totalVolume = 0
       Loop through rows in the data
       For j = 2 To RowCount
    *Get total volume for current ticker*

           If ActiveSheet.Cells(j, 1).Value = ticker Then
               'test = ActiveSheet.Cells(j, 8).Value
               totalVolume = totalVolume + ActiveSheet.Cells(j, 8).Value

           End If
           Worksheets(yearValue).Activate
    *Get starting price for current ticker*
           If ActiveSheet.Cells(j - 1, 1).Value <> ticker And ActiveSheet.Cells(j, 1).Value = ticker Then

               startingPrice = ActiveSheet.Cells(j, 6).Value

           End If

    *Get ending price for current ticker*
           If ActiveSheet.Cells(j + 1, 1).Value <> ticker And ActiveSheet.Cells(j, 1).Value = ticker Then

               endingPrice = ActiveSheet.Cells(j, 6).Value

           End If 


### Images showing table and runtime
#### VBA_Challenge 2017 before refactoring

! [VBA 2017] (/resources/VBA_Challenge_2017before.PNG)
#### VBA_CHallenge 2018 before refactoring
![VBA_Challenge 2018 with nested do loops.  The code took 0.84375 seconds to run.] (https://github.com/jcsargis00/stock-analysis/blob/main/resources/VBA_Challenge_2018before.PNG)
#### VBA_Challenge 2017 after refactoring
![VBA_Challenge 2017 with nested do loops.  The code took 0.125 seconds to run.] (https://github.com/jcsargis00/stock-analysis/blob/main/resources/VBA_Challenge_2017.PNG)
#### VBA_CHallenge 2018 after refactoring 
![VBA_Challenge 2017 with nested do loops.  The code took 0.125 seconds to run.] (https://github.com/jcsargis00/stock-analysis/blob/main/resources/VBA_Challenge_2018.PNG)
## Summary
### Advantages and Disadvantages of refactoring code
My approach to refactoring this code was to go after less nesting levels (do loops), reduce complexity in the conditional statements and reduce the line count.  Specific tasks included:
*  Reuse of much of the code written for the stock_analysis project in terms of the logic for determining how to track the volume, starting price and ending price for each ticker. 
* Reuse of code to make the report available for multiple years with a button.

### How do these pros and cons apply to refactoring the original VBA script.

