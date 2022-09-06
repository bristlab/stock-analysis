# Stock Analysis with Excel and VBA

## Overview of Project

In this project, we'll be using Excel and VBA to analyze stock data.

### Purpose

Our client has provided a dataset that includes historical information for a dozen stocks across two years. Previously, we wrote some custom tools using VBA which allowed us to quickly analyze the entire dataset and output performance metrics across all 12 stocks. We've decided to refactor our tools with the goal of increasing their efficiancy.

## Initial Analysis and Challenges

While our VB scripts perform well when applied to this dataset, we want to ensure that our solution will scale well when it comes to much larger datasets. To do this, we'll make some adjustments to *how* our code functions without changing the resulting output.

![2017 results before refactoring](https://raw.githubusercontent.com/bristlab/stock-analysis/main/Resources/VBA_Challenge_2017_original.png)


![2018 results before refactoring](https://raw.githubusercontent.com/bristlab/stock-analysis/main/Resources/VBA_Challenge_2018_original.png)

Before refactoring, our script analyzes the stock data for 2017 and 2018 in 0.58 seconds and 0.73 seconds respectively. The big challenge with refactoring this code is to incorporate better usage of arrays and conditional loops, while also keeping the operations consistent so as not to arrive at a different output.

Consider the following examples:

```
If Cells(j, 1).Value = ticker Then
   totalVolume = totalVolume + Cells(j, 8).Value
End If
```

becomes

```
If Cells(tV, 1).Value = tickers(tickerIndex) Then
    tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(tV, 8).Value
End If
```

While these two snippets of code do the same thing, the second snippet utilizes arrays and indexes to process the data more efficiently.

### Analysis After Refactoring


After refactoring our scripts, we are pleased to report that the output is identical to the previous version, with the only difference being that our new version executes faster. Our results for 2017 and 2018 were generated in 0.5 seconds and 0.52 seconds respectively.

![2017 results before refactoring](https://raw.githubusercontent.com/bristlab/stock-analysis/main/Resources/VBA_Challenge_2017_refactored.png)


![2018 results after refactoring](https://raw.githubusercontent.com/bristlab/stock-analysis/main/Resources/VBA_Challenge_2018_refactored.png)

While this improvement may not seem like much, we're satisfied with the results because the code is more scalable in that it's not only faster, but can also be adapted to larger datasets without significant retooling.

### Summary: Pros and Cons of Refactoring

Faster, more efficient code is always preferable, especially when working with very large datasets. In some cases, your code is nearly impossible to use because it runs far too slowly, and refactoring is the only way to make your code execute in a timely manner.

However, sometimes refactoring for efficiency comes at the cost of human readability and simplicity. While our new code is slightly faster and more flexible, it might not have been the best use of our time to spend several hours optimizing code to shave off a few tenths of a second. Another issue with refactoring is the potential to introduce new bugs in your code without realizing it, requiring additional time to be spent on troubleshooting.