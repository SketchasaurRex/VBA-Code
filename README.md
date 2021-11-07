# VBA-Code
## Overview of Project
### This project was designed to filter data on stocks by year to determine the Total Daily Volume and Return for a desired year organized by the Ticker. 

## Results

### This screenshot shows our initial run time when we first open Excel for the years of 2017 and 2018. This is a pretty high run time for a small ammount of code.

![alt text](https://github.com/SketchasaurRex/VBA-Code/blob/main/Resources/VBA_Challenge_2017.PNG) ![alt text](https://github.com/SketchasaurRex/VBA-Code/blob/main/Resources/VBA_Challenge_2018.PNG)

### Lets compare this time with our refactored code, under 0.1 seconds, incredibly faster. We were able to get the run times consistently below 0.08 seconds. It may not seem like much, but we will give a very strong real life application in our summary on just how significant this is.

![alt text](https://github.com/SketchasaurRex/VBA-Code/blob/main/Resources/Refactor_2017.png) ![alt text](https://github.com/SketchasaurRex/VBA-Code/blob/main/Resources/Refactor_2018.png)

### What we did to make the code faster is refractor the code, having three independent loops running. The nested loops, while thorough, added to the run times in comparison.

## Summary

### Refractoring the code reduces the time to run larger sets of data. Refractoring might take more time to write and execute but it takes less time to run. IF you're getting the desired results for your base code then you don't really need to refactor your code. The original code is condensed and uses less variables but will take more time to run, especially with larger data sets.
