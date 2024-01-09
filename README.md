# VBA-challenge
Module 2 Asignment

Sub StockData()

'Working though all workseets
Dim ws As Worksheet
For Each ws In ThisWorkbook.Worksheets

'Set a variable for holding the ticker name in column A
Dim Ticker_Name As String
    
'Set a varable for holding a total count on the total volume of trade
Dim Ticker_Total As Double
Ticker_Total = 0

'Keep track of the location for each ticker name in the summary table
Dim Summary_Table_Row As Integer
Summary_Table_Row = 2
        
'Note: Yearly Change is simply the difference: (Close Price at the end of a trading year - Open Price at the beginning of the trading year)
'Percent change is a simple percent change -->((Close - Open)/Open)*100
Dim open_price As Double

'Set initial open_price. Other opening prices will be determined in the conditional loop.
open_price = Cells(2, 3).Value
        
Dim close_price As Double
Dim yearly_change As Double
Dim percent_change As Double

'Label the row and columns headers
ws.range("I1").Value = "Ticker"
ws.range("J1").Value = "Yearly Change"
ws.range("K1").Value = "Percentage Change"
ws.range("L1").Value = "The Total Stock Volume"
ws.range("P1").Value = "Ticker"
ws.range("Q1").Value = "Value"
ws.range("O2").Value = "Greatest % Increase"
ws.range("O3").Value = "Greatest % Decrease"
ws.range("O4").Value = "Greatest Total Volume"


'Count the number of rows in the first column.
Lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

       For i = 2 To Lastrow

            'Searches for when the value of the next cell is different than that of the current cell
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
              'Set the ticker name
              Ticker_Name = ws.Cells(i, 1).Value
              
              'Print the Total Stock volume in summary table
              ws.range("L" & Summary_Table_Row).Value = Ticker_Total
              
             'Add the volume of trade
              Ticker_Total = Ticker_Total + ws.Cells(i, 7).Value

              'Now collect information about closing price
              close_price = ws.Cells(i, 6).Value

              'Calculate yearly change
              yearly_change = (close_price - open_price)
              
              'Print the yearly change for each ticker in the summary table
              ws.range("J" & Summary_Table_Row).Value = yearly_change

             'Check for the non-divisibilty condition when calculating the percent change
                If (open_price = 0) Then

                    percent_change = 0

                Else
                    
                    percent_change = yearly_change / open_price
                
            End If

              'Print the yearly change for each ticker in the summary table
              ws.range("K" & Summary_Table_Row).Value = percent_change
              ws.range("K" & Summary_Table_Row).NumberFormat = "0.00%"
   
              'Reset the row counter. Add one to the summary_ticker_row
              Summary_Table_Row = Summary_Table_Row + 1

              'Reset volume of trade to zero
              Ticker_Total = 0

              'Reset the opening price
              open_price = ws.Cells(i + 1, 3)
            
            Else
              
               'Add the volume of trade
             Ticker_Total = Ticker_Total + ws.Cells(i, 7).Value

            
        End If
        
    Next i
    
    
    
        
'Conditional formatting that will highlight positive change in green and negative change in red

lastrow_sum_table = ws.Cells(Rows.Count, 10).End(xlUp).Row
        
'Color code yearly change
    
        For j = 2 To lastrow_sum_table
               
            If ws.Cells(j, 10).Value > 0 Then
                    
                    ws.Cells(j, 10).Interior.ColorIndex = 10
                
                Else
                    
                    ws.Cells(j, 10).Interior.ColorIndex = 3
                    
            End If
        
        Next j
    
    

'testing variables to second summary table
                
        Test = WorksheetFunction.Max(ws.range("K2:K" & Summary_Table_Row))
        
        Test2 = WorksheetFunction.Match(Test, ws.range("K2:K" & Summary_Table_Row), 0)
        
        Test3 = ws.Cells(Test2 + 1, "I").Value
        
        ws.Cells(2, "P").Value = Test3
        
        ws.Cells(2, "Q").Value = Test
        
        ws.range("Q2" & Summary_Table_Row).NumberFormat = "0.00%"
        
        
        Test_1 = WorksheetFunction.Min(ws.range("K2:K" & Summary_Table_Row))
        
        Test_2 = WorksheetFunction.Match(Test_1, ws.range("K2:K" & Summary_Table_Row), 0)
        
        Test_3 = ws.Cells(Test_2 + 1, "I").Value
        
        ws.Cells(3, "P").Value = Test_3
        
        ws.Cells(3, "Q").Value = Test_1
        
        
        Test4 = WorksheetFunction.Max(ws.range("L2:L" & Summary_Table_Row))
        
        Test5 = WorksheetFunction.Match(Test4, ws.range("L2:L" & Summary_Table_Row), 0)
        
        Test6 = ws.Cells(Test5 + 1, "I").Value
        
        ws.Cells(4, "P").Value = Test6
        
        ws.Cells(4, "Q").Value = Test4
        
        ws.range("Q2:Q3" & Summary_Table_Row).NumberFormat = "0.00%"
        
 Next ws
 
End Sub
