# Stock Analysis Reports for Steve
### *Stock-analysis module 2 assignment*
## Project Overview: Purpose of this analysis
This project creates a report to analyze green stocks performance in 2017 and 2018.  The report is for Steve to advise his parents on investing in worthy "green" stocks.  The report was created with Excel, Visual Basic Application (VBA) and macros.  The initial report focused on stock DQ, a stock his parents were interested in. The final report analyzes an additional 11 "green" stock in terms of volume, starting price and ending price.
## Results
The first report produced results, but seemed to have room for improvement in terms of more elegant programming,
less lines of  codes and more efficient run times.  Although it executed and produced accurate report results, by refactoring the code, run times were reduced signficantly, whilst producing the same accurate results.
Refactoring included creation of arrays to reduce nested loop execution times.   For example, the inital code used two nested loops, going through the entire data sheet for each stock.  The second report used arrays, allowing the code to set up the tickers as an array and then loop through the data one time setting up the volume and % return for each stock.  The resulting code in the 2nd report ran over 6 times faster.

  
##Examples from the reports
# Images showing table and runtime
### VBA_Challenge 2017 before refactoring
![VBA 2017 with nested do loops took 0.84375 seconds to run.](resources/VBA_Challenge_2017before.PNG)
### VBA_CHallenge 2018 before refactoring
![VBA_Challenge 2018 with nested do loops took 0.84375 seconds to run.](resources/VBA_Challenge_2018before.PNG)
### VBA_Challenge 2017 after refactoring
![VBA_Challenge 2017 took 0.125 seconds to run.](resources/VBA_Challenge_2017.PNG)
### VBA_CHallenge 2018 after refactoring 
![VBA_Challenge 2017 took 0.125 seconds to run.](resources/VBA_Challenge_2018.PNG)
## Summary
### Advantages and Disadvantages of refactoring code
My approach to refactoring this code was to go after less nesting levels (do loops), reduce complexity in the conditional statements and reduce the code line count.  Specific tasks included:
*  Reuse of much of the code written for the stock_analysis project in terms of the logic for determining how to track the volume, starting price and ending price for each ticker. 
* Reuse of code to make the report available for multiple years with a macro button so Steve could run the report for either year 2017 or 2018.
* By altering existing code, it's pretty easy to break what was working if you don't pay attention to every detail and keep a copy of the working code available to compare and contrast changes.

### How do these pros and cons apply to refactoring the original VBA script.
Pros: More efficient reporting times, less coding.  Efficient run times will be even more important when dealing with very large data sets, especially during coding, if a mistake is introduced and the code has to be run multiple times to find the bug.
Cons: By setting up the tickers in an array to get rid of the nested loops, it was very easy to mix up the variables and introduce mistakes.  Additionally, by refactoring the code, even though there was less code, it was a little more complicated to follow how each variable was assigned.  I used the debug facilities to trace the variable assignments when the first reports of the refactored code volumes did not initially match the volumes of the first report.    


