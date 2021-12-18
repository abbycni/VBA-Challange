Attribute VB_Name = "Module2"
Sub Stock()
Cells(1, 8).Value = "Stock Ticker"
Cells(1, 9).Value = "Yearly Change"
Cells(1, 10).Value = "Percentage Change"
Cells(1, 11).Value = "Total Stock Volume"
Dim ticker As String

Dim Volume_Counter As Double
'Set starting row for volume counter
Volume_Counter = 0
Dim Summary_Table As Double
Summary_Table = 2

Dim Open_ticker As Double
Dim Yearly_percentage_change As Double
Dim Yearly_change As Double

Open_ticker = Cells(2, 3).Value

For R = 2 To 797711


If Cells(R + 1, 1).Value <> Cells(R, 1).Value Then

    
    
'Write Ticker

ticker = Cells(R, 1).Value
   
'Calculate  StocktVolume (Sum column G for all value in the ticker)
Volume_Counter = Volume_Counter + Cells(R, 7)
'Calculate yearly change (Closing value for the year- Opening value for the year)
Close_Price = Cells(R, 6).Value
Yearly_change = Close_Price - Open_ticker

    'Set 0 values to 0%

    If Yearly_change = 0 Then
    Yearly_percentage_change = 0
    ElseIf Open_ticker = 0 Then
    Yearly_percentage_change = 0
    

    Else
    'Calculate  yearly percentage change (YearlyChange/Opening Value)*100)
    Yearly_percentage_change = Yearly_change / Open_ticker * 100
    End If


'Print the Volume; counter per ticker
Range("k" & Summary_Table) = Volume_Counter

'Print the ticker name in the summary table
Range("h" & Summary_Table).Value = ticker
'Print yearly change in summary table

Range("i" & Summary_Table).Value = Yearly_change

'Print Yearly percentage change to Column j
Range("j" & Summary_Table).Value = Yearly_percentage_change

'Add one to the summary table row
Summary_Table = Summary_Table + 1

'Reset Volume_Counter
 Volume_Counter = 0
 Open_ticker = Cells(R + 1, 3).Value
 

'Add to Volume_Counter

Else

Volume_Counter = Volume_Counter + Cells(R, 7).Value

End If


Next R


'Conditional formating

For C = 2 To 3169
If Cells(C, 9).Value <= 0 Then
Cells(C, 9).Interior.ColorIndex = 3
    
    
    
    Else
    Cells(C, 9).Interior.ColorIndex = 4
End If

Next C

End Sub





