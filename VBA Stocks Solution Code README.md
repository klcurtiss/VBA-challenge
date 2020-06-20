# VBA-challenge
Sub MyVBAStocksHW()

Dim Ws As Variant

For Each Ws In Worksheets
    Dim WorksheetName As String
    Dim LastRow As Long
    Columns("A:P").EntireColumn.AutoFit
    Columns("B").NumberFormat = 0
    Columns("K").NumberFormat = "0.00%"
    
    'set up last row count for Column 1
    WorksheetName = Ws.Name
    LastRow = Ws.Cells(Rows.Count, 1).End(xlUp).Row
    LastColumn = Ws.Cells(1, Columns.Count).End(xlToLeft).Column
    
    'MsgBox "Worksheet " + WorksheetName
    
    'Set up new column titles
    Ws.Cells(1, 9) = "Ticker"
    'Ws.Cells(1, 10) = "Open Value"
    'Ws.Cells(1, 11) = "Close Value"
    Ws.Cells(1, 10) = "Yearly Change"
    Ws.Cells(1, 11) = "Percent Change"
    Ws.Cells(1, 12) = "Total Volume"
   
    Ws.Range("I1").Interior.ColorIndex = 17
    Ws.Range("J1").Interior.ColorIndex = 17
    Ws.Range("K1").Interior.ColorIndex = 17
    Ws.Range("L1").Interior.ColorIndex = 17
    'Ws.Range("M1").Interior.ColorIndex = 17
    
    'MsgBox "Titles Entered: Ticker, Yearly Change, Percent Change, and Total Stock Volume"
    
    Dim i As Long
    Dim j As Integer
    Dim k As Long
    Dim m As Long
    Dim v As Double

    j = 2
    k = 2
    m = 2
    v = 0
   
   For p = 2 To LastRow
    If Ws.Cells(p, 3) = 0 Then
        Ws.Cells(p, 3) = 1
        End If
        
   Next p
   
    For i = 2 To LastRow
    
         v = v + Ws.Cells(i, 7).Value
    
        'First, determine when the ticker symbols change in column A
         If Ws.Cells(i + 1, 1).Value <> Ws.Cells(i, 1).Value Then
            'Assign ticker symbol to column I
            Ws.Range("I" & j).Value = Ws.Cells(i, 1).Value
            'Assign yearly change to column J
            Ws.Range("J" & j).Value = Ws.Cells(i, 6) - Ws.Cells(m, 3)
            'Assign percent change to column K
            Ws.Range("K" & j).Value = (Ws.Range("J" & j).Value / Ws.Cells(m, 3).Value)
            'Assign the total volume to colum L
            Ws.Range("L" & j).Value = v
             j = j + 1
            m = i + 1
            v = 0
    
        End If
       
    Next i
    
    For r = 2 To LastRow
        If Ws.Cells(r, 10) > 0 Then
            Ws.Cells(r, 10).Interior.Color = vbGreen
            Else
            Ws.Cells(r, 10).Interior.Color = vbRed
        End If
    Next r
   
 Next Ws


End Sub
