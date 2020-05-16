Attribute VB_Name = "Module1"
'1. Output ticker symbol
'2. Change from opening price at beginning of year to end of year
'3. % change from above
'4. Total stock volume
'5. Color cells for + and -

Sub Alphatesting():

'Set an initial variable for Ticker name
Dim Ticker_Name As String

'Set an initial variable for total volume per ticker name
Dim Ticker_Total As Variant
Ticker_Total = 0

'Keep track of location for each Ticker
Dim Summary_Table_Row As Integer


'Set variables for Open/Close Stock Price for Yearly Change and % Change
Dim Stock_Open As Double
Dim Stock_Close As Double
Dim Yearly_Change As Double
Dim Unique_Value As Integer
Dim Percent_Change As Double


Dim LastRow As Long
Dim i As Long


    'Loop through all sheets
    For Each ws In Worksheets
    
     Summary_Table_Row = 2
        
        'Create column labels
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "% Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"

    'Determine the last row in each sheet
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        For i = 2 To LastRow
        
        'Check if we are in the same ticker name
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            'Establish unique Ticker Name -> return value to column I
            Ticker_Name = ws.Cells(i, 1).Value
            ws.Range("I" & Summary_Table_Row).Value = Ticker_Name
            
            'Add the total stock volume -> return value to column L
            Ticker_Total = Ticker_Total + ws.Cells(i, 7).Value
            ws.Range("L" & Summary_Table_Row).Value = Ticker_Total
            
            'Count the number of times each Ticker Name occurs in sheet
            Unique_Value = WorksheetFunction.CountIf(ws.Range("A:A"), Ticker_Name)
            
            'Define open/close stock values and compute change -> return value to column J
            Stock_Open = ws.Cells(i - (Unique_Value - 1), 3).Value
            Stock_Close = ws.Cells(i, 6).Value
            Yearly_Change = Stock_Close - Stock_Open
            ws.Range("J" & Summary_Table_Row).Value = Yearly_Change
            
            'Color code change
            If Yearly_Change > 0 Then
                ws.Range("J" & Summary_Table_Row).Interior.Color = vbGreen
            Else
                ws.Range("J" & Summary_Table_Row).Interior.Color = vbRed
            End If
                
            'Compute % change -> return value to column K
            If Stock_Open <> 0 Then
                Percent_Change = Yearly_Change / Stock_Open
            Else
                Percent_Change = 0
            End If
            ws.Range("K" & Summary_Table_Row).Value = Percent_Change
                        
            'Move to next row of table and reset Stock Volume
            Summary_Table_Row = Summary_Table_Row + 1
            Ticker_Total = 0
        
        'If the ticker cell following a row is the same
        Else
        Ticker_Total = Ticker_Total + ws.Cells(i, 7).Value
        
         End If
        Next i
    
    'Return Max % Increase, % Decrease and Total Volume
    ws.Range("Q2") = WorksheetFunction.Max(ws.Range("K:K"))
    ws.Range("Q3") = WorksheetFunction.Min(ws.Range("K:K"))
    ws.Range("Q4") = WorksheetFunction.Max(ws.Range("L:L"))
   
   'Match above values to ticker name
    ws.Range("P2") = WorksheetFunction.Index(ws.Range("I:I"), WorksheetFunction.Match(ws.Range("Q2"), ws.Range("K:K"), 0))
    ws.Range("P3") = WorksheetFunction.Index(ws.Range("I:I"), WorksheetFunction.Match(ws.Range("Q3"), ws.Range("K:K"), 0))
    ws.Range("P4") = WorksheetFunction.Index(ws.Range("I:I"), WorksheetFunction.Match(ws.Range("Q4"), ws.Range("L:L"), 0))
    
    'Autofit all columns and format % columns
    ws.UsedRange.Columns.AutoFit
    ws.Columns("K").NumberFormat = "0.00%"
    ws.Range("Q2", "Q3").NumberFormat = "0.00%"
                         
    Next ws

End Sub

