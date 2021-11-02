# stock-analysis

# Stock-Analysis

## Overview of Project
The main purpose of this analysis was to generate a summary of the total volume of activity and yearly return percentage for specified "green" stocks for the years 2017 and 2018.
A macro had already been developed to accomplish this objective, so I was tasked with refactoring the original code to increase efficiency and improve performance.

## Results

Results: Using images and examples of your code, compare the stock performance between 2017 and 2018, as well as the execution times of the original script and the refactored script.

In general, Green stocks performed much better in 2017 than they did in 2018. In fact, of the 12 stocks tracked, only one had a negative return in 2017 while in 2018 only 2 of the 12 stocks had a positive return.
DQ, the stock that was initially focused on, performed better than all the other stocks tracked in 2017. In 2018 however, DQ was the worst-performing stock and one of the least traded. I would recommend investing in ENPH
as it was heavily traded and had great positive returns in both 2017 and 2018. See the tables below for the Total Daily Volume and yearly return of the 12 green stocks tracked for 2017 and 2018.
<img src = "https://github.com/AaronAKTX/stock-analysis/blob/main/Resources/2017.PNG"> <img src = "https://github.com/AaronAKTX/stock-analysis/blob/main/Resources/2018.PNG"><img src = "https://github.com/AaronAKTX/stock-analysis/blob/main/Resources/2017.PNG",
The refactoring of the script to run the stock analysis was a great success. The speed the results are generated in is over 8 times faster than the original script. By changing the macro that the 'Run Analysis for All Stocks' button
is assigned to, it's easy to compare the old script run time to the new script run time. Below are some screenshots of the elapsed time of the run. The first two are when the original macro was run and the second two or from when the refactored macro was run.
<img src = "https://github.com/AaronAKTX/stock-analysis/blob/main/Resources/2017.PNG", width = "100"> <img src = "https://github.com/AaronAKTX/stock-analysis/blob/main/Resources/2018.PNG", width = "100">
#
#
<img src = "https://github.com/AaronAKTX/stock-analysis/blob/main/Resources/VBA_Challenge_2018.PNG", width = "100"> <img src = "https://github.com/AaronAKTX/stock-analysis/blob/main/Resources/2018_Original_Macro.PNG", width = "100">

The biggest gain from the refactoring was achieved by looping through the entire list of all the stocks starting and ending price and volume traded each day only once. By using arrays, I could check the ticker and capture and sum up the volume of trades and the starting and ending prices in one
loop rather than looping through the entire set one time for each specific ticker abbreviation.  I was also able to gain a little extra performance by moving the activate worksheet outside the For loop when setting the total volume and return on the All Stocks Analysis worksheet.

#Summary: In a summary statement, address the following questions.
What are the advantages or disadvantages of refactoring code?

There are a few advantages to refactoring code. One is performance improvement. This is a pretty big deal when results are time-critical. A poor-performing macro can take quite a while to produce results, and depending on the computer it can lock up resources while it's running and prohibit other tasks from being worked on.  Refactoring also gives a person a chance to look at existing code with fresh eyes and glean the intention of a macro, understand how it is supposed to work, and then apply updated or more efficient methodologies to improve performance and possibly fix previously unseen errors. It can also help reinforce best practices and improve a person's scriptwriting by giving them a starting place and a framework to build and improve.
I suppose a con with refactorization, especially if one is working with code that he/she didn't originally write, would be that the big picture of what the code was originally intended to do may not be easy to grasp right away. It's possible, especially with poorly commented code that more time is spent
trying to make sense of the code than actually improving it.

How do these pros and cons apply to refactoring the original VBA script?
The original VBA script we wrote during classwork was great. It returned the correct results and helped us practice writing for loops and get a feel for working with data. The refactoring of the script was very helpful because it brought to light inefficiencies that I didn't
even realize were there. It helped me realize the original macro was cycling through the rows of data 12 times and helped lead me to a more efficient solution of looping through one time and using arrays to hold values for multiple stocks without having to start back at the beginning for each stock.



