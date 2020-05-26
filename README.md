# stock-analysis

# Challenge

Refractored te original data - PLEASE USE TEST SHEET FOR GRADING (I wanted to compare and contrast against original code.)


The code places tickers, volume, starting prices, and ending prices in arrays:

    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
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


If more time was afforded, would have liked to figure out how to solve for ticker names without manually placing them into the VBA sub.

TickerIndex is set to zero, and begins to loop through line items in the selected year's spreadsheet. The loop does 4 things, checks the value of the ticker, increases the count of the ticker, and depending on the value of the ticker's symbol and of it's the first or last time it shows up, effects the starting and ending prices in Test sheet Analysis. 

Formatting was slightly edited to output results into test sheet. 
