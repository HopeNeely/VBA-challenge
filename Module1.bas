Attribute VB_Name = "Module1"
Sub AlphabeticalTesting()

'Bonus:Loop through all sheets... After figuring out market days per year function...
'AND add "ws." into all cells listed in code
'For Each ws In Worksheets

'Set column headers in Summary Table
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"
'Set row and column titles for Bonus Summary Value table
'Cells(2, 15).Value = "Greatest % Increase"
'Cells(3, 15).Value = "Greatest % Decrease"
'Cells(4, 15).Value = "Greatest Total Volume"
'Cells(1, 16).Value = "Ticker"
'Cells(1, 17).Value = "Value"

'Create initial variable to hold the ticker symbol
Dim ticker_symbol As String

'Create initial variable to hold the yearly change in price from begining open to closing end prices
Dim yearly_change As Double

'Create initial variable to hold begining open price
Dim begining_open_price As Double

'Create initial variable to hold closing price at year end
Dim end_closing_price As Double

'Create initial variable to hold the percent change in price
Dim percent_change As Double

'Create variable to hold greatest percent increase for bonus table
Dim max_increase As Double

'Create variable to hold greatest percent decrease for bonus table
Dim max_decrease As Double

'Create inital variable to hold the sum of stock volume
Dim total_volume As Double
total_volume = 0

'Create variable for the greatest total volume for bonus table
Dim max_volume As Double

'Keep track of the location for each ticker symbol in summary table
Dim Summary_Table_Row As Integer
Summary_Table_Row = 2

' Determine the Last Row in stock data
LastRow = Cells(Rows.Count, 1).End(xlUp).Row

'Create initial begining opening price
begining_open_price = Cells(2, 3).Value

'Loop through all stock info
For i = 2 To LastRow

'Check if we are still within the same ticker symbol, if not...
If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

'Set the ticker symbol
ticker_symbol = Cells(i, 1).Value

'Set the end closing price
end_closing_price = Cells(i, 6).Value


'Subtract begining opening price from end closing price (Broken on Worksheet P)
yearly_change = end_closing_price - begining_open_price

'Find Percent change from begining opening price to closing price at year end
'Divide the yearly change by begining opening price
    If begining_open_price > 0 Then
    percent_change = yearly_change / begining_open_price
    Else
    percent_change = 1

    End If

'Add total volume of the stock
total_volume = total_volume + Cells(i, 7).Value

'Print the ticker symbol in summary table row
Range("I" & Summary_Table_Row).Value = ticker_symbol

'Print the total volume in summary table row
Range("L" & Summary_Table_Row).Value = total_volume

'Print yearly change in summary table row
Range("J" & Summary_Table_Row).Value = yearly_change

   'Set conditional formatting that will highlight positive yearly change value green
    If yearly_change > 0 Then
    Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
    
    'Highlight negative yearly change values red
     ElseIf yearly_change < 0 Then
     Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
     
     End If

'Print percent change in summary table row
Range("K" & Summary_Table_Row).Value = percent_change

'Add one to the summary table row
Summary_Table_Row = Summary_Table_Row + 1

'Reset the total volume
total_volume = 0

'Find Yearly Change from begining opening price to closing price at year end
'Set the begining opening price
begining_open_price = Cells((i + 1), 3).Value

'If the cell immediately following a row shows the same ticker...
Else

'Add to the total_volume
total_volume = total_volume + Cells(i, 3).Value

End If

Next i

' Determine the Last Row in summary table
'LastRow = Cells(Rows.Count, 11).End(xlUp).Row

'Loop through all summary table data
'For j = 2 To LastRow

'Establish greatest percent increase for bonus table
'max_increase = Application.WorksheetFunction.Max(Cells(j, 11))

'Establish greatest percent decrease for bonus table
'max_decrease = Application.WorksheetFunction.Min(Cells(j, 11))

'Establish greatest total stock volume for bonus table
'max_volume = Application.WorksheetFunction.Max(Cells(j, 12))


'Check if percent change is greatest percent increase then...
'If Cells(j, 11).Value = max_increase Then

    'Retieve ticker symbol associated with greatest percent increase and print in bonus table
    'Cells(2, 16) = Cells(j, 9).Value
    'Cells(2, 17) = max_increase
    
'Check if percent change is greatest percent decrease then...
'ElseIf Cells(j, 11).Value = max_decrease Then

    'Retrieve ticker symbol associated with greatest percent decrease and print in bonus table
    'Cells(3, 16) = Cells(j, 9).Value
    'Cells(3, 17) = max_decrease
    
'Check if Total Stock Volume is greatest total volume then...
'ElseIf Cells(j, 12).Value = max_volume Then
    
    'Retrieve ticker associated with greatest total volume and print in bonus table
    'Cells(4, 16) = Cells(j, 9).Value
    'Cells(4, 17) = max_volume
    
'End If

'Next j

'For bonus finish worksheet loop by sending to next ws and setting finishing message
'Next ws
'MsgBox ("Stock Summary Complete")

End Sub
