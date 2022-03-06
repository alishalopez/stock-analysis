# Stock Analysis
## Overview of Project
Steve was looking for a simple and efficient way to analyze numerous stocks in a dataset. His parents were first interested in the stock “DQ”, but after research and seeing they were at a negative annual return for 2018 at about 63%, it was clear that may not be a smart investment. If Steve was able to do the same for each stock, it would save him time and allow for comparison between all of them. Through refactoring code in Excel by VBA, I was able to analyze the entire stock market and give him results on key performance indicators that may tell him if stocks were successful or not over the last few years so he was able to make quick decisions and not have to decipher lines of data on his own.

## Results
To display the results of the code, I only used three header titles as the combination of the three was sufficient enough to give Steve a better understanding of how each stock did at a glance for that particular year after the user entered into the input box set. The headers were Ticker, Total Daily Volume and the Annual Return. 
```VBA
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"
```
For loops were essential to this code, as setting formulas so it knows when to go to the next ticker and calculate the starting and ending prices were key parts. For example, instead of manually entering starting prices, we set j to start at 2 and go through all of the rows so we could use an if statement to check that the row above had a different ticker value so that we knew it was the start of another. 
```VBA
 If Cells(j - 1, 1).Value <> tickers(tickerIndex) Then
        tickerEndingPrices(tickerIndex) = Cells(j, 6).Value

tickerIndex = tickerIndex + 1
            
End If
```
To find the ending price, we used that same code, but had it set for Cells (j + 1, 1).

Again I used for loops to enter the results under the correct headers for easy reading. 
```VBA
Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
```
Formatting the data was also entered into our code to keep it clean and organized, including changing the interior color to whether Annual Return was negative or positive just by looking at the color shown. 

From 2017 to 2018, there was a drastic change in the Annual Returns. In 2017, there was only one that was negative as opposed to 2018 when there was only 2 out of 12 that was positive. That raises a lot of questions as to the market or if there was a specific incident that affected them all. Steve is able to dig deeper as he finds suitable. 

![2017 Stocks](https://github.com/alishalopez/stock-analysis/blob/0c32d1f8c2a0126f5ef8c18ddc081181c28189e7/resources/2017%20Stocks.png)

![2018 Stocks](https://github.com/alishalopez/stock-analysis/blob/0c32d1f8c2a0126f5ef8c18ddc081181c28189e7/resources/2018%20Stocks.png)

Refactoring the code allows for the code to run much faster. With the original script I had created, it was taking over a second to run as compared to now at the fraction of the time. 

![VBA_Challenge_2017](https://github.com/alishalopez/stock-analysis/blob/0c32d1f8c2a0126f5ef8c18ddc081181c28189e7/resources/VBA_Challenge_2017.png)
![VBA_Challenge_2018](https://github.com/alishalopez/stock-analysis/blob/0c32d1f8c2a0126f5ef8c18ddc081181c28189e7/resources/VBA_Challenge_2018.png)

## Summary
There are many advantages of refactoring code such as reducing run time and doing more at once. I was able to reduce steps and clean up the code for easy reading and understanding of what was happening on each line. Refactoring the code allowed for me to take out repeating steps I may have had in the original and really understand what I was trying to accomplish in this project. Using less lines I would have overall if I ran each stock one by one makes for a more logical approach.

A huge disadvantage of refactoring code in my opinion is the amount of time it may take to rewrite the code for it to accomplish what you have in mind as it has to be set to all variables you are trying to solve for. You have to be very tedious to ensure each line is able to accomplish your goal. The more knowledge you have in refactoring code, the easier it will be, but that comes with more practice and experience. 

I thought I had it all figured out in the original VBA script to later find a more efficient way that I could do more with the refactored. I can now quickly change the refactored if more stocks were added or some were taken out. The original may have been easier and less time consuming, but overall I would rather spend more time on the refactored and have learned more about the data to use later on. 
