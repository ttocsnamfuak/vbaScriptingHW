Sub wsStockChanges()
    
    'Setup and initalize all my variables
    Dim TickerVol As Double
    TickerVol = 0
    Dim Ticker As String
    Dim NbrStocks As Long
    NbrStocks = 0
    Dim LastRow As Long
    LastRow = 0
    Dim ResultLastRow As Long
    ResultLastRow = 0
    Dim RowCounter As Integer
    RowCounter = 0
    Dim i As Long
    i = 0
    Dim StockOpen As Double
    StockOpen = 0
    Dim StockClose As Double
    StockClose = 0
    Dim StockPerChg As Double
    StockPerChg = 0
    Dim StockYrChg As Double
    StockYrChg = 0
    Dim OpenCounter As Integer
    OpenCounter = 1


    For Each ws In Worksheets
        'How many rows to loop through
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

        'Setup Colums for totals
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        RowCounter = 2

        ' Loop through rows in the column
        For i = 2 To LastRow
        ' For i = 2 To 22 - test code
            ' Searches for when the value of the next cell is different than that of the current cell
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

                ' Prepare to print the stock totals in the next available row
                Ticker = ws.Cells(i, 1).Value
                ws.Cells(RowCounter, 9).Value = Ticker

                TickerVol = TickerVol + ws.Cells(i, 7).Value
                ws.Cells(RowCounter, 12).Value = TickerVol

                StockClose = ws.Cells(i, 6).Value
                'MsgBox "Close/OPen" StockClose & StockOpen
                StockYrChg = (StockClose - StockOpen)
                ws.Cells(RowCounter, 10).Value = StockYrChg

                StockPerChg = (1 - (StockOpen / StockClose))
                ws.Cells(RowCounter, 11).Value = StockPerChg

                'Reset variables for next Stock & increase the Display Row
                TickerVol = 0
                StockClose = 0
                StockOpen = 0
                StockYrChg = 0
                StockPerChg = 0
                OpenCounter = 1
                RowCounter = RowCounter + 1
            
            ElseIf OpenCounter = 1 Then
                TickerVol = TickerVol + ws.Cells(i, 7).Value
                'set stock opening value
                StockOpen = ws.Cells(i, 3).Value
                'MsgBox "Open" & StockOpen
                OpenCounter = OpenCounter + 1
            Else
                TickerVol = TickerVol + ws.Cells(i, 7).Value
            End If
            
        Next i

    'Format the cells Green for positive year change and red for negative
        'How many rows to loop through
        ResultLastRow = ws.Cells(Rows.Count, 10).End(xlUp).Row
        For y = 2 To ResultLastRow
                If ws.Cells(y, 10).Value <= 0 Then
                    ws.Cells(y, 10).Interior.ColorIndex = 3
                Else
                    ws.Cells(y, 10).Interior.ColorIndex = 4
                End If
        Next y

    'Determine Larget %, Smallest % and highest volume
    
    
    'Setup Rows for Greatest, Least, etc
       ' ws.Cells(1, 16).Value = "Ticker"
        'ws.Cells(1, 17).Value = "Volume"
        'ws.Cells(2, 15).Value = "Greatest % Increase"
        'ws.Cells(3, 15).Value = "Greatest % Decrease"
        'ws.Cells(4, 15).Value = "Greatest Volume"


        'Dim rng As Range
        'Dim dblMin As Double

        'Set range from which to determine smallest value
        'Range("A6", Cells(10,LastColumn))
        'Set rng = ws.Range("K2", "K" & ResultLastRow)


        'Worksheet function MIN returns the smallest value in a range 

        'dblMin = Application.WorksheetFunction.Min(rng)
        'ws.Cells(3, 16).Value = dblMin
        'Displays smallest value
        'MsgBox dblMin


   Next ws

End Sub
