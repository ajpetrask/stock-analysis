Stock Analysis Using VBA

Project Overview 

A friend asked me to help his parents look into a few stocks to see if they were worth purchasing. I set-up an Excel spreadsheet and utilized VBA to compare the stocks. After his parents looked it over they could easily see which ones were worth the investment.

Purpose

The purpose of this project was to simplifiy the decision making process when it comes to investing in single stocks. 

Results

To refactor the code, I first copied the code that need be reworked. Then I followed the preset steps that set the stucture for refactoring. 

Refactored Code

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
    Worksheets(yearValue).Activate
    
    'Get the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    '1a) Create a ticker Index
        tickerIndex = 0
        

    '1b) Create three output arrays
        Dim tickerVolumes(12) As Long
        Dim tickerStartingPrices(12) As Single
       ''2a) Create a for loop to initialize the tickerVolumes to zero.
        For i = 0 To 11
            tickerVolumes(i) = 0
            tickerStartingPrices(i) = 0
            tickerEndingPrices(i) = 0
        Next i
    
    '2b) Loop over all the rows in the spreadsheet.
        For i = 2 To RowCount
        
         '3a) Increase volume for current ticker
            tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
               If Cells(i - 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(ickersindex) Then
               tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            End If  Dim tickerEndingPrices(12) As Single
        
            '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        'If  Then
            If Cells(i + 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(ickerIndex) Then
                tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
                
            End If

            '3d Increase the tickerIndex.
            If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i = 1, 1).Value <> tickers(tickerIndex) Then
            tickerIndex = tickerindx + 1
            
            End If
    
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
    
    Worksheets("All Stocks Analysis").Activate
        Cells(i + 4, 1).Value = tickers(tickersIndex)
        Cells(i + 4, 2).Value = tickerVolumes(tickerIndex)
        Cells(i + 4, 3).Value = tickerEndingPrices(tickerIndex) / tickerStartingPrice
        
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
    MsgBox "This code ran in " & (endTime - startTime) & " sec"

End Sub

2017
<img width="355" alt="2017" src="https://user-images.githubusercontent.com/101996888/163291925-eaf12a56-5b99-4477-8f84-76b2cde802d2.png">


2018
<img width="355" alt="2018" src="https://user-images.githubusercontent.com/101996888/163291967-2fd5c21e-273b-4cc6-8bb5-1897cd17be1d.png">

Summary

Refactoring Code

Refactoring code can make code easier cleaner and easier to read. it can also make it run more efficient. I did not find any real disadvantages to refactoring code other than incorecctly refactoring, rendering the code useless.

Refactoring VBA Script
    
I think a major advantage to using VBA is being able to utilize multiple modules at once to code by comparision. The disadvantage I found was using syntax.    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
