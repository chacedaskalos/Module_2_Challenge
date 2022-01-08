# Module_2_Challenge

##**Overview**
  ###The purpose of this anlysis is to "refactor" the subroutine I initially used to help Steve analyze stocks. I refactored the subroutine to make it run more efficiently. This is important because if Steve were to want to analyze a larger datset he could do it faster with the new code.

##**Results**
  ###The results of the refactored script are phenomenal. Anlysis on 2017 stocks with the starter code ran in 0.89 seconds, while analysis on 2017 stocks with the improved code ran in 0.18 seconds
![Starter code run_time 2017](https://user-images.githubusercontent.com/96211484/148657215-02b1f0c4-b244-472c-bf20-9b4e59a0a675.png)
![VBA_Challenge_2017](https://user-images.githubusercontent.com/96211484/148657221-fa8017a7-9599-4424-8140-156aed58a909.png)
Success with year 2018 was similar as well. 2018 analysis starter code ran in 0.96 seconds while 2018 analysis with improved code ran in 0.16 seconds. 
![Starter code  run_time 2018](https://user-images.githubusercontent.com/96211484/148657336-b4470c0a-4b4a-48f0-a930-5590e33260eb.png)
![VBA_Challenge_2018](https://user-images.githubusercontent.com/96211484/148657337-09afad71-62ea-40e1-8b67-26ae65bf22cc.png)
That is a whopping 394% faster for 2017 and 500% faster for 2018. The most important change to the code was the use of arrays for the output data.

```
        For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        
        tickerIndex = i
        
        Cells(4 + i, 1).Value = tickers(tickerIndex)
        Cells(4 + i, 2).Value = tickerVolumes(tickerIndex)
        Cells(4 + i, 3).Value = tickerEndingPrices(tickerIndex) / tickerStartingPrices(tickerIndex) - 1
        
         Next i 
 ```
 ##Summary ###The biggest advantage to refactoring code in general is that it can be more efficient, cleaner, and take up less memory. The biggest disadvantage is the time it takes to refactor. Sometimes the subroutine that is written isn't the best but it gets the job done. It can be tidious to go through the code and find how to improve it.
 
 For this challenge specifically the advantage was the new code obviously ran faster. However, as humans we can barely even tell the difference between 0.18 seconds and 0.89 seconds, so on this scale it might be more tidious to refactor the code than the benefit of the saved ~0.70 seconds.
    
    
