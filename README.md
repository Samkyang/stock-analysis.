# Challenge 02: Stock-Analysis

## 1. Overview of Project

#### The overview of this project was to help Steve research certain stocks from 2017 and 2018 so that he can provide feedback to his parents. It was also to learn how to refract somebody elses code to make it run more efficient as well as how to leverage code I already wrote from going through this weeks reading.

### How It Was Done

#### I created a macro that would scan the stock lists from 2018 and 2017. The macro would take the user input and output the ticker symbol, total daily volume for the given year and the annual return. 

#### When reviewing the macro from _*Download challenge_starter_code.vbs*_ I reviewed the provided macro and challenge requirements and saw that it would be fairly similar to the what I had just completed in the reading.

## 2. Results

### Stock Performance
#### The performance of the stocks all generally all made less returns than they did in the previous year except for RUN and TERP which did better than they did than the previous year.

![YoY Comparison](https://github.com/Samkyang/stock-analysis./blob/main/YoY%20Comparison.png)

### Execution times of script
#### The refracting of the script made a huge difference. When I first had run the script it took 53469.68 seconds and 53443.6 seconds for 2017 and 2018 respectively. When I refracted the script it took .3164062 seconds and .3203125 seconds for 2017 and 2018 respectively. This is a huge time difference just by writting the lines of code to be more efficient.

![Green stocks 2017](https://github.com/Samkyang/stock-analysis./blob/main/Green%20Stocks%202017.png)
##### Initial code for 2017
![Challenge 2017](https://github.com/Samkyang/stock-analysis./blob/main/Challenge%202017.png)
##### refracted code for 2017

![Green stocks 2018](https://github.com/Samkyang/stock-analysis./blob/main/Green%20Stock%202018.png)
##### Initial code for 2018
![Challenge 2018](https://github.com/Samkyang/stock-analysis./blob/main/Challenge%202018.png)
##### refracted code for 2018

How I was able to calculate the time was by using the Timer Function then displaying it in a message box.

``startTime = Timer``

``...``

``EndTimer = Timer``

``MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)``

#### More Useful codes were nested ``for`` statements.
#### The way these were used were by putting a `for` statement inside another `for` statement so that I could have the macro run through all the whole dataset and find specific tickers, sum the volume based off of ticker, as well as find the starting and ending price, 

    '1a) Create a ticker Index
   
    For j = 0 To 11
    tickerIndex = Tickers(j)
   
    '1b) Create three output arrays
       
        Dim tickerStartingPrices As Single
        Dim tickerEndingPrices As Single
        Dim tickerVolumes As Long

    '2a) Create a for loop to initialize the tickerVolumes to zero
    '2b) Loop over all the rows in the spreadsheet.
        Sheets(yearValue).Activate
        tickerVolumes = 0
        For i = rowStart To RowCount

       
        '3a) Increase volume for current ticker
       
             If Cells(i, 1).Value = tickerIndex Then

                'store and add to volume
                tickerVolumes = tickerVolumes + Cells(i, 8).Value
               
             End If
       
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
           
        If Cells(i, 1).Value = tickerIndex And Cells(i - 1, 1).Value <> tickerIndex Then
         
         'store starting price
         tickerStartingPrices = Cells(i, 6).Value
       
         End If
           
       
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        'If  Then
           
           If Cells(i, 1).Value = tickerIndex And Cells(i + 1, 1).Value <> tickerIndex Then
           
            'store ending price
            tickerEndingPrices = Cells(i, 6).Value

        '3d Increase the tickerIndex.
           
           
        End If
   
   
    Next i
   
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
       
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + j, 1).Value = tickerIndex
        Cells(4 + j, 2).Value = tickerVolumes
        Cells(4 + j, 3).Value = tickerEndingPrices / tickerStartingPrices - 1

    Next j


## 3. Summary
#### Pros and Cons of refractoring code
#### Pros
##### The benefits of refractoring code is that you don't have to start from scratch and write everything yourself. You can leverage what was already written or how the format wants to be displayed. Comments could also have been left in the script to outline what their objective is. 

##### Initially when I downloaded the script I saw that there was already an outline as to what was needed to be accomplised and saw the formatting requirements. Seeing this gave me a better understanding of how to structure my script.

    '1a) Create a ticker Index
    

    '1b) Create three output arrays   
    
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero. 
    
        
    ''2b) Loop over all the rows in the spreadsheet. 
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
            
            
            
        'End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        'If  Then
            
            

#### Cons
##### The drawbacks of refactoring code is that you sometimes don't understand why some code is written a certain way or if they have *Magic Numbers* instead of setting values. For instance if the data set changed and we now start code on line 3 and we didn't use the `StartRow = 2` line of code and just left it as `2`, then the next person coming in to review my script might be confused why we are grab data from line 2.
