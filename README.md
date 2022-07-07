# Stock Analysis Using VBA

## Overview of Project
This project utilises different data points from 12 different stocks over the span of two years (2017 and 2018). The data points include: stock ticker, date, open price, highest daily price, lowest daily price, closing price, adjusted closing price, and daily volume. These data can be used to analyse stock returns, market interest and possibly future price movements.

### Purpose
The purpose of the project is to create a **VBA script** that can automate the process of computing each stock's volume and return for a given year. Starting with an initial script, the goal was to refactor it and optimize the runtime. By doing so, Steve will be able to help his parents in a better way by analysing many more stocks in less time than before.

## Results
### Overview
The results have been overwhelmingly positive. 
For the year 2017, the runtime was decreased by an impressive 84.3%. This percentage decreased was measured only one time. For a more accurate result, it is imperative to run each script multiple times, average the runtimes and compute the difference. <br>
![image](https://user-images.githubusercontent.com/85131345/177840996-555a82be-f6cf-471a-9106-49c010d3bf1d.png)
![image](https://user-images.githubusercontent.com/85131345/177841045-e8148d8e-d0f5-48bc-8bbb-c85254086083.png)
<br>
For the year 2018, the runtime was decreased by 80.7%. <br>
![image](https://user-images.githubusercontent.com/85131345/177841441-dfa20875-0566-479b-8316-c623c8c4b6b8.png)
![image](https://user-images.githubusercontent.com/85131345/177841468-37377134-4706-47b7-b455-133bec7548e5.png) <br> <br>
### Explanation
As for the reasons that generated this optimization, there are several that will be explained with the following code and reasoning:
- Inefficient looping over all the sheets' rows:
  - The original script looped over all the rows twelve times looking for a specific ticker while the new script only looped once over all the rows and stored all the data for each ticker.
  ```
  Initial Script
  
  For i = 0 To 11
    
        ticker = tickers(i)
        totalVolume = 0
        Worksheets(yearValue).Activate
        
        For j = 2 To RowCount
            If Cells(j, 1).Value = ticker Then
  ```
- Variables in the refactored script were declared differently
  - The starting and ending price were declared as Double type while on the refactored code they are Single.
  - As for the volume traded, in the initial code, it was not explicitly declared while on the refactored one it is declared as a Long type.
  ```
    Initial Script
    
    Dim startingPrice As Double
    Dim endingPrice As Double
    
    Refactored Script
    
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
  ```
 - Starting Price, Ending Price, and Total Volume were converted into arrays
    - In the initial script, the only array was the ticker one and the other three main variables were changed in each loop. In the refactored code, these variables were converted into arrays to be able to efficiently loop through them. In the code from the previous section, it is possible to grasp the difference in the declaration of variables.
## Summary
### Code Refactoring Advantages and Disadvantages
Refactoring code certainly has many advantages, one of which has been clearly shown in this project. However, not everything is easily done. The following will show in a detailed manner the advantages and disadvantages of code refactoring:
#### Advantages
- Refactoring can optimize runtime in an impressive manner
  - A script might work perfectly but it does not mean it is efficient
- Refactoring will make the code more readable, either by cleaning the script or by adding comments
  - Unreadable code will make it hard to work across teams
  - Bugs will be identified and removed during this process
#### Disadvantages
- Time consuming
  - It is possible that the original code was written by another person and getting accustomed to it might take a while
- Sometimes it is just better to start from the beginning
  - Very poorly written code will be quite hard to refactor and it might just be easier to start from scratch
 - Poorly refactored code will create errors
 ### Original VBA Script VS Refactored Script
 #### Advantages
 - Exponentially faster runtime of new script
  - A roughly 80% lower time
 - Better documentation
  - This will make the script easier to mantain and understand in the future
  - The refactored script can now be used for analysing thousands of stocks
 #### Disadvantages
 - There are no significant disadvantages of the refactored code, at least in technical terms. The only disadvantages could be for the people who were already accustomed to the original script. Other than that, the code is much more efficient, scalabale and mantainable now. 
