# stock-analysis

## Overview of Project

### Purpose

##### Steve recently graduated in Finance. His parents wanted to support him and become his first customers and invest in green energy. Without researching, his parents decided they wanted to invest all their money into a stock known as DAQO. Steve wanted to diversify their funds and so he came to us to analyze stock options for his parent's investment. 

##### We created a workbook for him that allowed him to analyze an entire dataset at the click of a button. Now, in order to do more research for his parents, Steve wanted us to expand the dataset to include the entire stock market over the last few years. 

##### In order to provide Steve with everything he needed to assist his parent's in their investment strategies, we refactored the workbook to loop through all the data one time in order to collect the same information, faster. 

## Results

##### First, I copied the code from the original workbook into a new Module, in order to create an input box, chart headers, ticker array, and to activate the necessary worksheet. The refactored code has been provided below, with instructions written within. The purpose of writting instruction comments within the code, is to make it easier to understand and more easily readable.  

```
Sub AllStocksAnalysisRefactored()
    Dim startTime As Single
    Dim endTime  As Single

    yearValue = InputBox("What year would you like to run the analysis on?")

    startTime = Timer
    
    'Format the output sheet on All Stocks Analysis worksheet
    Worksheets("All Stocks Analysis").Activate
    
    Range("A1").Value = "All Stocks (" + yearValue + ")"
    
    'Create a header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"

    'Initialize array of all tickers
    Dim tickers(12) As String
    
    tickers(0) = "AY"
    tickers(1) = "CSIQ"
    tickers(2) = "DQ"
    tickers(3) = "ENPH"
    tickers(4) = "FSLR"
    tickers(5) = "HASI"
    tickers(6) = "JKS"
    tickers(7) = "RUN"
    tickers(8) = "SEDG"
    tickers(9) = "SPWR"
    tickers(10) = "TERP"
    tickers(11) = "VSLR"
    
    'Activate data worksheet

    Sheets(yearValue).Activate
    
    'Get the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    '1a) Create a ticker Index
    
    tickerIndex = 0

    '1b) Create three output arrays
    
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    
    For i = 0 To 11
    
        tickerVolumes(i) = 0
        tickerStartingPrices(i) = 0
        tickerEndingPrices(i) = 0
        
    Next i
        
        
    ''2b) Loop over all the rows in the spreadsheet.
    
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
            
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
        
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        
        End If
            
        'End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        'If  Then
            
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
         
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            
        End If

        '3d) Increase the tickerIndex.
            
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
        
            tickerIndex = tickerIndex + 1
            
        End If
            
        'End If
    
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
        
    Next i
    
    'Formatting
    Worksheets("All Stocks Analysis").Activate
    Range("A3:C3").Font.FontStyle = "Bold"
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "#,##0"
    Range("C4:C15").NumberFormat = "0.0%"
    Columns("B").AutoFit

    dataRowStart = 4
    dataRowEnd = 15

    For i = dataRowStart To dataRowEnd
        
        If Cells(i, 3) > 0 Then
            
            Cells(i, 3).Interior.Color = vbGreen
            
        Else
        
            Cells(i, 3).Interior.Color = vbRed
            
        End If
        
    Next i
 
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

End Sub
```

## Summary

### Advantages of Refactoring Code

##### Refactoring code is a key part of the coding process. By refactoring, we are able to make our code more efficient - by taking fewer steps, using less memory, and improving the logic of the code to make it easier for future users to read. 

### Disadvantages of Refactoring Code

##### Although refactoring is a key process, it has its risks. Improper refactoring can sometimes lead to new bugs and errors developing within your code. In addition, if a group of people are all working together to refactor code, coordination is imperative. Otherwise, the code could become less neat, which could set the refactoring process back a few steps. 

### Advantages and Disadvantages of the Refactored VBA Script

##### This was a good chance to understand refactoring and how it can decrease the macro run time. In our original analysis for Steve, ti took about a second for the code to run. After refactoring, the code ran in ~0.20 seconds. Meaning, refactoring provided us our analysis around five times quicker than the original code. The main disadvantage of refactoring this analysis, is that refactoring is for benefit to the viewer. Although the macro runs and provides outputs faster now, there is extra work that goes into refactoring code which could be argued as a delay in results for the analysis. Overall; however, refactoring seems to be of benefit to both the user, and future users. 

![VBA_Challenge_2017](https://user-images.githubusercontent.com/103145291/175141724-7fbd4896-916b-4b72-bb41-ac535d088be0.PNG)

![VBA_Challenge_2018](https://user-images.githubusercontent.com/103145291/175141742-92ebb390-b3ec-4e11-913c-6c93308c1681.PNG)


