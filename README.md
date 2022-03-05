# VBA of Wall Street

## Overview of Project 

### To utilize Microsoft Visual Basic for Applications (VBA) with the intent to refactor a code’s efficiency so that it may be more applicable for analysis of the entire stock market. 

## Refactoring VBA

### 1a) Ticker Index 

Starting with Challenge Starter Code from Module 2, I first sought to establish the ``tickerIndex`` variable as equal to 0 (as almost all programming languages start at zero). This was also necessary as it was later used to access the correct indexes between four arrays: the ``tickers`` array, the ``tickerVolumes`` array, the ``tickerStartingPrices`` array, and the ``tickerendingPrices`` array. And so, I entered the following code into VBA:

    tickerIndex = 0
  
### 1b) Output Arrays

After creating the ticker index, I then created three additional output arrays. The ``tickerVolumes`` array was created as a Long data type. While the ``tickerStartingPrices`` and ``tickerEndingPrices`` were created as Single data types. The additions to the code were as follows: 

    Dim tickerVolumes(12) As Long
  
    Dim tickerStartingPrices(12) As Single
  
    Dim tickerEndingPrices(12) As Single

Twelve was inserted into the parenthesis as that is the number of elements (stocks) we are currently concerned with. 

### 2a) TickerVolumes Loop

Next, I needed to create a ``for`` loop that would initialize the ``tickerVolumes`` to zero. In addition, I also needed to specify that the loop would iterate twelve times, once for each element (stock). As a result, the added code looked like this: 

    For i = 0 To 11
       tickerVolumes(i) = 0 
    
### 2b) Looping Over All Rows

But I also needed to ensure that the loop iterated over every row. To do this I needed my loop to start at 2, avoiding the header row, and continuing until the rows were fully counted. This finished piece of code looked like the following:

    For i = 2 To RowCount

### 3a) Increase TickerVolumes

After writing that loop, I needed to write an additional script that would add the current ticker volumes to the current stock ticker. This script would use the ``tickerIndex`` as the variable for the index. Knowing that ‘Volume’ was in the eighth row, the script looked at follows: 

    tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value

### 3b) First Row ‘If-Then’ 

The next task was creating an ``if-then`` statement that would determine if the current row was the first row, while using the ``tickerIndex``. If that statement proved to be true, then the script would assign the current starting price to the ``tickerStartingPrices`` variable. As a result, I wrote the following script: 

    If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex)
  
    Then tickerStartingPrices(tickerIndex) = Cells(i, 6).Value

The first half of the ``if`` script (before the ``And``) checks whether the current row’s ticker is within the ``tickerIndex`` while the second half checks whether the previous row’s ticker is not. If both statements are deemed true, then the ``tickerStartingPrice`` is assigned. 

### 3c) Last Row ‘If-Then’ 

Creating a ``if-then`` statement to determine whether the current row was the last row, was markedly similar to the ‘First Row Check.’ However ``Cells(i-1,1).Value`` is substituted with ``Cells(i+1, 1).Value`` to account for the fact that we are concerned with the row below are current row, rather than above. As a result, the statement looked like this: 

    If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex)
  
    Then tickerEndingPrices(tickerIndex) = Cells(i, 6).Value

### 3d) Increase TickerIndex

The last ``if-then`` statement sought to increase the ``tickerIndex`` if the next row’s ticker and the previous row’s ticker did not match. To do this I wrote the following script: 

    If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex)
  
    Then tickerIndex = tickerIndex + 1

The script specifies that if the next row’s ticker ``(Cells(i + 1, 1)`` does not equal the ``tickerIndex``, then the ``tickerIndex`` will be increased by one. 

### 4) ‘For’ Loop Output

Finally, I created a ``for`` loop that ran through all four of my arrays, ``tickers``, ``tickerVolumes``, ``tickerStartingPrices``, and ``tickerEndingPrices``, and output them to ‘Ticker’, ‘Total Daily Volume and ‘Return’ respectively. Since that header row was placed in the third row, I knew these values needed to start at the fourth row, descending downward until all loops were completed. The resulting script is as follows: 
    
    Worksheets("All Stocks Analysis").Activate
    
        Cells(4 + i, 1).Value = tickers(i)
        
        Cells(4 + i, 2).Value = tickerVolumes(i)
        
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1

## Summary

### General Advantages & Disadvantages

One advantage to refactoring VBA code is that it can help create a more streamlined script that other programmers can more readily understand and use right away. Similarly, if the script is more streamlined, it could be less prone to human error. Finally, if the script is easy to use by programmers who had no part in writing the code and it remains error-free, it will undoubtedly save its users time and ultimately money. 

Yet, there are some disadvantages to refactoring code. If you are refactoring a lengthy script, doing so could be a time-consuming endeavor. You could also invertedly delete or alter a portion of previously workable code rendering the new script unusable or inaccurate. 

### Advantages & Disadvantages of Refactoring this VBA Script

Regarding the refactored script, the advantages I found aligned with those purported in the previous section. The script runs the code at a quicker rate. This fact would be extremely important if used on the Stock Market, where buying and selling of stocks rapidly remains of preeminent importance. This speedy load time can be seen by the images that follow. 

![VBA 2017 Loading Screen](https://github.com/chrisknox97/stock-analysis/blob/main/Resources/VBA_Challenge_2017.png) 
![VBA 2018 Loading Screen](https://github.com/chrisknox97/stock-analysis/blob/main/Resources/VBA_Challenge_2018.png)

While the disadvantages of the original VBA script stemmed from its lack of concise programming language, resulting in me receiving error messages more often than I would have preferred. I can also imagine more streamlined VBA as less accessible to newcomers to the field. 

