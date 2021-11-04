# VBA-Code
## Overview of Project
### This project was designed to filter data on stocks by year to determine the Total Daily Volume and Return for a desired year organized by Ticker. In learning how to do this with nested for loops we noticed that the code run time seemed to take more time than we would like. This led to us refactoring our code to get faster results.

## Results


### This screenshot shows our initial run time when we first open Excel for the year 2018.

alt text alt text

The two screenshot's above show the run times for our original code at around 0.65 seconds. This is a pretty high run time for a small ammount of code. Lets compare this time with our refactored code!

alt text

Here is our inital run time for the refactored code. Under 0.1 seconds! this is much faster!

alt text alt text

We were able to get the run times consistently below 0.08 seconds. It may not seem like much, but we will give a very strong real life application in our summary on just how significant this is.

So what did we do to run the code faster? In our original code we had nested for loops which added to the run time. When we were looping through the tickers we then looped through the rows of data using the ticker to grab our volume and prices. This method caused us to output our code within the nested loop as well, most likely increasing run time.

Our code looked a bit like this

'''

For i = 0 To 11
    ticker = tickers(i)
    totalVolume = 0
    
    Worksheets(yearValue).Activate
    For j = 2 To RowCount
        
        If Cells(j, 1).Value = ticker Then
        
            totalVolume = totalVolume + Cells(j, 8).Value
        
        End If'
        
        If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
            
            startingPrice = Cells(j, 6).Value'
        
        End If'
        
        If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
            
            endingPrice = Cells(j, 6).Value'
            
        End If
        
    Next j
    
    Worksheets("All Stocks Analysis").Activate'
    Cells(4 + i, 1).Value = ticker'
    Cells(4 + i, 2).Value = totalVolume'
    Cells(4 + i, 3).Value = (endingPrice / startingPrice) - 1
    
Next i
'''

While this worked and allowed us to grab all the data with multiple if statements it caused multiple changing values simultaneously.

When refractored, we instead had three loops running independently. One for the tickerVolume, another to loop over the rows, and the last to output the data.

Refactored our code looks like this instead. While it is written longer (had to define more variables), the computation takes less time to execute.

'''

For j = 0 To 11
    tickerVolumes(j) = 0
Next j
    
For i = 2 To RowCount

    tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
    
    
    If Cells(i - 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
    
        tickerStartingPrices(tickerIndex) = Cells(i, 6)
    End If
        
    If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1) = tickers(tickerIndex) Then
    
        tickerEndingPrices(tickerIndex) = Cells(i, 6)
       
        tickerIndex = tickerIndex + 1
        
    End If

Next i

For i = 0 To 11
    
    Worksheets("All Stocks Analysis").Activate
    Cells(4 + i, 1).Value = tickers(i)
    Cells(4 + i, 2).Value = tickerVolumes(i)
    Cells(4 + i, 3).Value = (tickerEndingPrices(i) / tickerStartingPrices(i)) - 1
    
Next i
'''

Summary
If you do not need to refactor the code and you are getting desired results then refrain from doing so. The code might take a bit longer to execute, but it takes less time to write and is condensed. Testing run times on data load is important if you plan on using it for bigger sets of data down the road. If you have a high variation in the size of data, test it on a bigger set and refactor as needed.

In our VBA code, like stated above, it took longer to code via refactoring. Our original code was condensed and used less variables. The plus side to our refactored code is now we can apply much larger sets of data for the tickers to go over.
