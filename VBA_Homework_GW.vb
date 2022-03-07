'## Instructions

'Create a script that loops through all the stocks for one year and outputs the following information:

  '* The ticker symbol.

  '* Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.

  '* The percent change from opening price at the beginning of a given year to the closing price at the end of that year.

  '* The total stock volume of the stock.

'**Note:** Make sure to use conditional formatting that will highlight positive change in green and negative change in red.

'## Bonus

'Add functionality to your script to return the stock with the "Greatest % increase", "Greatest % decrease",
'and "Greatest total volume".

'Make the appropriate adjustments to your VBA script to allow it to run on every worksheet (that is, every year)
'just by running the VBA script once.

'## Other Considerations

'* Use the sheet `alphabetical_testing.xlsx` while developing your code. This data set is smaller and will
'allow you to test faster. Your code should run on this file in less than 3 to 5 minutes.

'* Make sure that the script acts the same on every sheet. The joy of VBA is that it takes the tediousness
'out of repetitive tasks with one click of a button.



Sub wall_street()


'Declare the dimensions

Dim yearly_change As Double
Dim percent_change As Double
Dim total_stock_vol As Long

Dim daily_change As Double
Dim average_change As Double

Dim start As Double
Dim i As Long


Dim lastrow As Long

Dim ws As Worksheet

Dim GreatestIncrease As Double
Dim GreatestDecrease As Double


'loop through the worksheets
    
    For Each ws In Worksheets
        ' Set values for each worksheet
        j = 0
        total_stock_volume = 0
        yearly_change = 0
        start = 2
        daily_change = 0
        
        ' Set column header
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("p1").Value = "Ticker"
        ws.Range("q1").Value = "Value"
        ws.Range("o2").Value = "Greatest % Increase"
        ws.Range("o3").Value = "Greatest % Decrease"
        ws.Range("o4").Value = "Greatest Total Volume"
             
    
        ' find and record the last row with data
        lastrow = Cells(Rows.Count, "A").End(xlUp).Row

            For i = 2 To lastrow
            
                'when the ticker changes store the results in the variable called total_stock_volume
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

                    total_stock_volume = total_stock_volume + ws.Cells(i, 7).Value
                
                    'look after the rows with zero total_stock_volumes by printing these values
                
                    If total_stock_volume = 0 Then
                    
                    ws.Range("I" & 2 + j).Value = Cells(i, 1).Value
                    ws.Range("J" & 2 + j).Value = 0
                    ws.Range("K" & 2 + j).Value = "%" & 0
                    ws.Range("L" & 2 + j).Value = 0
                
                Else
                    ' find non_zero starting values for the calculation
                    If ws.Cells(start, 3) = 0 Then
                        For non_zero = start To i
                            If ws.Cells(non_zero, 3).Value <> 0 Then
                                start = non_zero
                                Exit For
                            End If
                        Next non_zero
                    End If
                    
                    ' calculate the change between the ticker start value and the final ticker value
                    daily_change = (ws.Cells(i, 6) - ws.Cells(start, 3))
                    percent_change = Round((daily_change / ws.Cells(start, 3) * 100), 2)

                    ' move on the the next ticker
                    start = i + 1

                    ' print the to columns created in each work sheet
                    
                    ws.Range("I" & 2 + j).Value = ws.Cells(i, 1).Value
                    ws.Range("J" & 2 + j).Value = Round(daily_change, 2)
                    ws.Range("K" & 2 + j).Value = "%" & percent_change
                    ws.Range("L" & 2 + j).Value = total_stock_volume
                    
                    ' conditional formatting highlighting positive change in green and negative change in red
                    
                    If percent_change > 0 Then
                    ws.Range("J" & 2 + j).Interior.ColorIndex = 4
                    
                    ElseIf percent_change < 0 Then
                    ws.Range("J" & 2 + j).Interior.ColorIndex = 3
                    
                    Else
                    ws.Range("J" & 2 + j).Interior.ColorIndex = 0
                    
                    End If
                                  
                                       
                
                End If

                ' reset the values for the next ticker
                
                total_stock_volume = 0
                yearly_change = 0
                j = j + 1
                daily_change = 0
                
            ' If the ticker hasnt changed print the result at the end of the row
            Else
                total_stock_volume = total_stock_volume + ws.Cells(i, 7).Value

            End If
            
        Next i
                
       ' find the "Greatest % increase", "Greatest % decrease" & "Greatest total volume" for each sheet and populate into the correct cells
       
       
       GreatestIncrease = Application.WorksheetFunction.Max(ws.Columns("K")) * 100
       GreatestDecrease = Application.WorksheetFunction.Min(ws.Columns("K")) * 100
       
       
       
        
       ws.Range("p2").Value = "Ticker1"
       ws.Range("q2").Value = "%" & GreatestIncrease
       ws.Range("p3").Value = "Ticker2"
       ws.Range("q3").Value = "%" & GreatestDecrease
       ws.Range("p4").Value = "Ticker3"
       ws.Range("q4").Value = Application.WorksheetFunction.Max(ws.Columns("l"))
       
       
             
       ' Autofit column width to suit data and make the font Bold
       
       ws.Columns("A:q").AutoFit
       ws.Range("I1:q1").Font.Bold = True
       ws.Range("o2:o4").Font.Bold = True
       
       
        
              
    Next ws
    
    
    

End Sub



