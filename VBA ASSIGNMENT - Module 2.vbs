VBA ASSIGNMENT - Module 2

Sub Stock_Analysis()

'Run Process on all worksheets
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        ws.Activate

'---------------------------------------------------------

    'Set the variables
    Dim Ticker As String
    Dim Ticker_Volume As Double
    Dim Ticker_Open As Double
    Dim Ticker_Close As Double
    Dim Yearly_Change As Double
    Dim Percent_Change As Double
    Dim LastRow As Long
    Dim i As Long
    Dim Per_Increase As Double
    Dim Per_Decrease As Double
    Dim Tot_Vol As Double
    Dim Ticker_Increase As Double
    Dim Ticker_Decrease As Double
    Dim Ticker_Total As Double
       
    'Set Title Row
    
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Quarterly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    Range("O1").Value = "Ticker"
    Range("P1").Value = "Value"
    Range("N2").Value = "Greatest % Icrease"
    Range("N3").Value = "Greatest % Decrease"
    Range("N4").Value = "Greatest Total Volume"
    
    'Set Initial Values
    Ticker_Volume = 0
    Ticker_Open = 0
    Ticker_Close = 0
    Yearly_Change = 0
    Percent_Change = 0
    
    'Set Variable for Last Row
    
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row

    'Location for Storing Ticker Name in the summary table
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2

    'Loop through all line items
    For i = 2 To LastRow

        If Cells(i - 1, 1).Value <> Cells(i, 1).Value Then

            'Set Ticker Open
            Ticker_Open = Cells(i, 3).Value
            
            'Set Ticker Volume
            Ticker_Volume = Ticker_Volume + Cells(i, 7).Value

        'Check if value is within the same Ticker Name
        ElseIf Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            
            'Set the Ticker Name
            Ticker = Cells(i, 1).Value

            'Set Ticker Close
            Ticker_Close = Ticker_Close + Cells(i, 6).Value
      
            'Set the Yearly Change
            Yearly_Change = Ticker_Close - Ticker_Open
      
            'Set the Percent Change
            Percent_Change = (((Ticker_Close - Ticker_Open) / Ticker_Open))

            'Add to the Ticker Volume
            Ticker_Volume = Ticker_Volume + Cells(i, 7).Value

            'Print the Summary Table
            Range("I" & Summary_Table_Row).Value = Ticker
            Range("J" & Summary_Table_Row).Value = Yearly_Change
            Range("K" & Summary_Table_Row).Value = Percent_Change
            Range("L" & Summary_Table_Row).Value = Ticker_Volume

            'Add one to the summary table row
            Summary_Table_Row = Summary_Table_Row + 1
      
            'Reset the totals
            Ticker_Close = 0
            Ticker_Volume = 0

        Else

            'Add to the Ticker Volume
            Ticker_Volume = Ticker_Volume + Cells(i, 7).Value
        End If
    Next i

'---------------------------------------------------------

    'Format positive and negative Quarterly Changes
    Dim quarterlyChange As Range
    Set quarterlyChange = Range("J:J")
    
    For Each Cell In quarterlyChange
        If Cell.Value < 0 Then
            Cell.Interior.ColorIndex = 3
        ElseIf Cell.Value > 0 Then
            Cell.Interior.ColorIndex = 4
        Else
            Cell.Interior.ColorIndex = 0
        End If
    Next

    'Format percent change as percentage
    Range("K:K").NumberFormat = "0.00%"
    
'---------------------------------------------------------

    'Create New Table with the highest and lowest
    
    
    'Set range to determine searches
    Search_Percent = Range("K:K")
    Search_Volume = Range("L:L")

    'Use min max to find the values
    Per_Increase = Application.WorksheetFunction.Max(Search_Percent)
        Cells(2, 16).Value = Per_Increase
        
    Per_Decrease = Application.WorksheetFunction.Min(Search_Percent)
        Cells(3, 16).Value = Per_Decrease
        
    Tot_Vol = Application.WorksheetFunction.Max(Search_Volume)
        Cells(4, 16).Value = Tot_Vol

    'Use match function to find the ticker names
    
    Ticker_Increase = WorksheetFunction.Match(WorksheetFunction.Max(Range("K2:K" & LastRow)), Range("K2:K" & LastRow), 0)
        Range("O2") = Cells(Ticker_Increase + 1, 9)

    Ticker_Decrease = WorksheetFunction.Match(WorksheetFunction.Min(Range("K2:K" & LastRow)), Range("K2:K" & LastRow), 0)
        Range("O3") = Cells(Ticker_Decrease + 1, 9)

    Ticker_Total = WorksheetFunction.Match(WorksheetFunction.Max(Range("L2:L" & LastRow)), Range("L2:L" & LastRow), 0)
        Range("O4") = Cells(Ticker_Total + 1, 9)
    
    'Format Percents
    Range("P2:P3").NumberFormat = "0.00%"
    ws.Columns("A:Z").AutoFit
    
'---------------------------------------------------------
    
'Loop process on all worksheets
Next ws
    
End Sub

