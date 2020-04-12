# Stock-Analysis
 Module 2 VBA
# Challenge
## Refactor the Analysis code with four different arrays using tickerIndex
Sub AllStocksAnalysis()
    yearValue = InputBox("What year would you like to run the analysis on?")

    Worksheets("All Stocks Analysis").Activate

    Range("A1").Value = "All Stocks (" + yearValue + ")"

    'Create a header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"
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


    Worksheets(yearValue).Activate

    'get the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    Dim totalVolume(12) As Single
    Dim startingPrice(12) As Double
    Dim endingPrice(12) As Double
    
    
    
    tickerIndex = 0
    
    
    ' total Volume array with all 12 elements
    For i = 0 To 11

        totalVolume(i) = 0
    
        
        Worksheets(yearValue).Activate

        'loop over all the rows
        For j = 2 To RowCount
    
            
            If Cells(j, 1).Value = tickers(i) Then

            'increase totalVolume by the value in the current row
            totalVolume(i) = totalVolume(i) + Cells(j, 8).Value
        
     
            End If
            
            
        Next j
        
        
    Next i
    

    ' startingPrice array with all 12 elements
    For k = 0 To 11

        startingPrice(k) = 0
        
        Worksheets(yearValue).Activate

        'loop over all the rows
        For l = 2 To RowCount
    
            
            If Cells(l - 1, 1).Value <> tickers(k) And Cells(l, 1).Value = tickers(k) Then

                startingPrice(k) = Cells(l, 6).Value

            End If
            
        Next l

    Next k
    
    tickerIndex = 0
    ' endingPrice array with all 12 elements
    For m = 0 To 11

        endingPrice(m) = 0

        
        Worksheets(yearValue).Activate

        'loop over all the rows
        For n = 2 To RowCount
    

            If Cells(n + 1, 1).Value <> tickers(m) And Cells(n, 1).Value = tickers(m) Then

                endingPrice(m) = Cells(n, 6).Value

            End If
            
        Next n

    Next m
     
     
     
    tickerIndex = 0
    'access all four arrays and fill in the sheet
    For p = 0 To 11
            

        Worksheets("All Stocks Analysis").Activate
        
        Cells(4 + p, 1).Value = tickers(tickerIndex)
        Cells(4 + p, 2).Value = totalVolume(tickerIndex)
        Cells(4 + p, 3).Value = endingPrice(tickerIndex) / startingPrice(tickerIndex) - 1
        
    
        If Cells(5 + p, 1).Value <> tickers(tickerIndex) Then
        tickerIndex = tickerIndex + 1
                
                
    End If


    Next p
     
 
 
 
 
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

    
End Sub
