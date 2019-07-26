Sub StockReview():

'Loop through sheets
For Each ws In Worksheets
    ws.Activate

    'Declare Variable
    Dim Vol_tot As Double
    Dim chr As Integer 'This variable will push the next stock the next line on my table
    Dim YearOpen As Double
    Dim YearClose As Double
    Dim YearChange As Double
    Dim YearPerChange As Double
    Dim MaxPerChange As String
    Dim MinPerChange As String
    Dim MaxVol_tot As String
    Dim MaxName As String
    Dim MaxPerTicker As String
    Dim MinPerTicker As String
    Dim MaxVolTicker As String
    Dim MaxSoFar As Double
    Dim MinSoFar As Double
    Dim MaxVolTotSoFar As Double
    Dim Last_Row_Table As Double
    Dim Last_Row As Double
    
    
    
    'Determine last row of the data
    Last_Row = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    'Set the row for 2 for the table
    chr = 2
    
    'Set up tables
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    
    'Making font bold for ease of reading
    ws.Cells(1, 9).Font.Bold = True
    ws.Cells(1, 10).Font.Bold = True
    ws.Cells(1, 11).Font.Bold = True
    ws.Cells(1, 12).Font.Bold = True
    ws.Cells(1, 16).Font.Bold = True
    ws.Cells(1, 17).Font.Bold = True
    ws.Cells(2, 15).Font.Bold = True
    ws.Cells(3, 15).Font.Bold = True
    ws.Cells(4, 15).Font.Bold = True

    'Adding a bottom boarder
    ws.Cells(1, 9).Borders(xlEdgeBottom).LineStyle = xlContinuous
    ws.Cells(1, 10).Borders(xlEdgeBottom).LineStyle = xlContinuous
    ws.Cells(1, 11).Borders(xlEdgeBottom).LineStyle = xlContinuous
    ws.Cells(1, 12).Borders(xlEdgeBottom).LineStyle = xlContinuous
    ws.Cells(1, 16).Borders(xlEdgeBottom).LineStyle = xlContinuous
    ws.Cells(1, 17).Borders(xlEdgeBottom).LineStyle = xlContinuous
    
    'Loop through all stocks
    For i = 2 To Last_Row
    
        'Check if its not the same stock
        If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
            Vol_tot = Vol_tot + ws.Cells(i, 7)
            ws.Cells(chr, 9).Value = ws.Cells(i, 1)
            ws.Cells(chr, 12).Value = Vol_tot
            YearClose = ws.Cells(i, 6).Value
            
            
            'Caluculate the year change, enter it into table and change format to dollars
            YearChange = YearClose - YearOpen
            ws.Cells(chr, 10).Value = YearChange
            
            'Conditionally Formatting Yearly Change
            If ws.Cells(chr, 10).Value > 0 Then
                ws.Cells(chr, 10).Interior.ColorIndex = 4
            ElseIf ws.Cells(chr, 10).Value < 0 Then
                ws.Cells(chr, 10).Interior.ColorIndex = 3
            Else
                ws.Cells(chr, 10).Interior.ColorIndex = 2
            End If
                
            
            'Calculate the percentage change, enter it into table
            If YearOpen > 0 Then
                YearPerChange = (YearClose - YearOpen) / YearOpen
                ws.Cells(chr, 11).Value = YearPerChange
            Else
                ws.Cells(chr, 11).Value = 0
            End If
                
            'Reset the vol total to zero
            Vol_tot = 0
            
            'Move variable to the next line for the table result
            chr = chr + 1
            
        'Checking to see if its the first date of the Year
        ElseIf ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
            YearOpen = ws.Cells(i, 3).Value
        
        Else
            Vol_tot = Vol_tot + ws.Cells(i, 7)
        
        End If
    
    Next i
    
    'Getting the Ticker symbol for the Max, Min for PerChange
    'Set the variables to zero
    MaxSoFar = 0
    MinSoFar = 0
    MaxVolTotSoFar = 0
    
    'Loops to find the Tickers of the Greatest increase and decrease
    ws.Activate
    'Determine last row of the table
    Last_Row_Table = ws.Cells(Rows.Count, 9).End(xlUp).Row
    
    'Building the table for Hard section of the problem
    'Getting the % Increse and % Decrease
    For i = 2 To Last_Row_Table
        If ws.Cells(i, 11).Value > MaxSoFar Then
            MaxSoFar = ws.Cells(i, 11).Value
            MaxPerTicker = ws.Cells(i, 9).Value
        ElseIf ws.Cells(i, 11).Value < MinSoFar Then
            MinSoFar = ws.Cells(i, 11).Value
            MinPerTicker = ws.Cells(i, 9).Value
        Else
        End If
    Next i
    
    'Getting the Ticker symbol and the value for Max Volume
    For i = 2 To Last_Row_Table
        If ws.Cells(i, 12).Value > MaxVolTotSoFar Then
            MaxVolTotSoFar = ws.Cells(i, 12).Value
            MaxVolTicker = ws.Cells(i, 9).Value
        Else
        End If
    Next i
    
    'Populate the Tickers into the table
    ws.Range("P2").Value = MaxPerTicker
    ws.Range("P3").Value = MinPerTicker
    ws.Range("P4").Value = MaxVolTicker
    'Populate the values in the table
    ws.Range("Q2").Value = MaxSoFar
    ws.Range("Q3").Value = MinSoFar
    ws.Range("Q4").Value = MaxVolTotSoFar

    
    'Autofitting columns of table to make column labels readable
    ws.Range("I:L").EntireColumn.AutoFit
    ws.Range("O:Q").EntireColumn.AutoFit
    'Format numbers of the table
    ws.Range("J2:J" & Last_Row_Table).NumberFormat = "$#,##0.0000000"
    ws.Range("K2:K" & Last_Row_Table).NumberFormat = "#0.00%"
    ws.Range("L2:L" & Last_Row_Table).NumberFormat = "###,###,###,##0"
    ws.Range("Q2:Q3").NumberFormat = "0.00%"
    ws.Range("Q4").NumberFormat = "###,###,###,##0"
    ws.Range("K2:K" & Last_Row_Table).NumberFormat = "0.00%"
    
    
Next ws
            
End Sub
