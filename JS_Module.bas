Attribute VB_Name = "Module1"
Sub TotalStockCalc_easy()
'Defining my variables
'In order to create loop through each worksheet
Dim ws As Worksheet
Dim starting_ws As Worksheet
'Loop through each worksheet
For Each ws In ThisWorkbook.Worksheets
    ws.Activate

'TotalStock_Vol will be used as counter for adding stock
Dim TotalStock_Vol As Double
    TotalStock_Vol = 0
'so we can find last row of sheet without manually inputting
Dim lastrow As Long
lastrow = Cells(Rows.Count, 1).End(xlUp).Row
'Create a variable to display ticker names once, instead of as many times exist in the sheet
Dim Ticker As String
'TotalStock_Display will display the totalstock_vol, before it resets
Dim TotalStock_Display As Long


'Setting up titles for TotalStock_Display column, Ticker column & Percent Change

Range("I1").value = "Ticker"
Range("J1").value = "Total Stock Volume"

'so that we can display each ticker with its associated total stock, set up a counter. When we add one one to the counter, we move down one row per ticker

Dim Summary_Table As Integer
'row 2 is where we want to start displaying numbers
Summary_Table = 2

'set up loop and if/then
    For i = 2 To lastrow
        If Cells(i + 1, 1).value <> Cells(i, 1).value Then
            'add each column total stock value
            TotalStock_Vol = TotalStock_Vol + Cells(i, 7).value
            'enter ticker for table
            Ticker = Cells(i, 1).value
            'display ticker
            Cells(Summary_Table, 9).value = Ticker
            'display total stock for that ticker
            Cells(Summary_Table, 10).value = TotalStock_Vol
            'add row to the table
            Summary_Table = Summary_Table + 1
            'reset total stock so we can count total stock again for new ticker
            TotalStock_Vol = 0
        Else
        TotalStock_Vol = TotalStock_Vol + Cells(i, 7).value
        End If
    Next i
    
'because of worksheet loop

Next


End Sub
