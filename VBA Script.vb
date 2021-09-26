Sub TickerChallenge()
'define variables

Dim ws As Worksheet
'loop Worksheet
For Each ws In ActiveWorkbook.Worksheets
ws.Activate

'define values

    Dim ticker As String
    Dim volume As Double
    Dim open_value As Double
    Dim close_value As Double
    Dim yearly_change As Double
    Dim percent_change As Double

'set headers
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"


'set values for summary
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2

   
'define starting volume

    volume = 0
    open_value = Cells(2, 3).Value
 
'set up loop to fill table
    For i = 2 To 70926
    
     If Cells(i - 1, 1) = Cells(i, 1) And Cells(i + 1, 1) <> Cells(i, 1).Value Then
    
    'yearly change
    close_value = Cells(i, 6).Value
    yearly_change = close_value - open_value
    Range("J" & Summary_Table_Row).Value = yearly_change
    
    'percent change
    If (open_value = 0 And close_value = 0) Then
    percent_change = 0
    ElseIf (open_value = 0 And close_value <> 0) Then
    percent_change = 1
    Else
    percent_change = yearly_change / open_value
    Range("K" & Summary_Table_Row).Value = percent_change
    End If
    
    'ticker
    ticker = Cells(i, 1).Value
    Range("I" & Summary_Table_Row).Value = ticker
        
    'total volume
    volume = volume + Cells(i, 7).Value
    
    Range("L" & Summary_Table_Row).Value = volume
            
    'set data to move rows
    volume = 0
    Summary_Table_Row = Summary_Table_Row + 1
    
    Else
        volume = volume + Cells(i, 7).Value
    End If
    Next i
    
Next ws
End Sub

