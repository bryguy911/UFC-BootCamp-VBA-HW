VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Sub ABC_Brian()
' Sort the stocks as pre-sorted

    ' LOOP by rows

    ' Make leaderboard/ tables

    ' IF next Stock is different, that means we have finished our group

    ' ELSE, then keep summing the Volume
    
    ' Was given a memory issue so simple solutaion was to change long to longlong( yeah that makes total sense)
'Stocks variable
    Dim ws As Worksheet
    Dim ticker As String
    Dim stock As String
    Dim next_stock As String
    Dim volume As LongLong
    Dim volume_total As LongLong
    Dim i As Long
    Dim leaderboard_row As Long
    Dim lastRow As Integer
    Dim column As LongLong
    
   
    'Precent change variables
    Dim open_price As Double
    Dim closing_price As Double
    Dim change As Double
    Dim pct_change As Double
    
    For Each ws In ThisWorkbook.Worksheets
    
    'worksheet headers 1
    ws.Range("t1").Value = "ticker"
    ws.Range("j1").Value = "Quarterly Change"
    ws.Range("k1").Value = "precent Change"
    ws.Range("l1").Value = "total stock volume"
    ' Reset the stock total , open price
    volume_total = 0
    open_price = Cells(2, 3).Value
    
    leaderboard_row = 2
    lastRow = ws.Cells(Rows.Count, 1).End(xlUp).row
    
    ' If statment for the stock value in first table
    For i = 2 To lastRow
        ' extract values from workbook, like the stocks, volume,next stocks
        stock = ws.Cells(i, 1).Value
        volume = ws.Cells(i, 7).Value
        closing_price = ws.Cells(i, 6).Value
        next_stock = ws.Cells(i + 1, 1).Value

        ' if statement
        If (stock <> next_stock) Then
            ' add total volume_total + volume
            volume_total = volume_total + volume
            'Math statements with defined variables
            change = closing_price - open_price
            closing_price = Cells(i, 6).Value
            pct_change = change / open_price * 100
            
            ' write to leaderboard(Table1)
            ws.Cells(leaderboard_row, 12).Value = volume_total
            ws.Cells(leaderboard_row, 11).Value = change
            ws.Cells(leaderboard_row, 10).Value = pct_change
            ws.Cells(leaderboard_row, 9).Value = stock
       
       
       'Conditional Formating- must be done in the loop for it to work 4=green
       If (change > 0) Then
       ws.Cells(leaderboard_row, 10).Interior.ColorIndex = 4
       ' Else if as we want it to turn red but dont turn 0 red leave white
       ElseIf (change < 0) Then
        ws.Cells(leaderboard_row, 10).Interior.ColorIndex = 3
        Else
        End If
       
        ' reset total
           volume_total = 0
            leaderboard_row = leaderboard_row + 1
            open_price = ws.Cells(i + 1, 3).Value
        Else
            ' add total
            volume_total = volume_total + volume
        End If
        Next i
        ' I dont know why i cant make it to work again after I looped the sheet :(
        
        End Sub
    
    
       
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
 
