# VBA-Stock-Analysis


' This subroutine loops through all worksheets in a workbook of stock data to calculate the yearly change,
' percent change and total volume traded for each ticker symbol found.  One row of output is generated per
' ticker symbol.  Yearly change cell is highlighted in red if there is a loss for the year, or green if
' even or there is a gain.  Percent change is formatted as a percent.
'
' Greatest volume traded, greatest percent increase, and greates percent decrease are tracked and output
' for each worksheet
'
' stockAnalyzer() assumes that the worksheet data is sorted by ticker symbol, and date
'

  - stockAnalyzer.bas  - Visual Basic script to parse Stock Data, calculating stock price change over time, percent change over time, and total volume traded.  Script assumes
                          worksheets are sorted by ticker, and by date.
  - stockAnalyzer.txt  - Text verstion of Visual Basic script.
  
  - 2014Results.png    - Screenshots of the top of the output results of each worksheet processed.
    2015Results.png
    2016Results.png

![image](/2016Results.png)
