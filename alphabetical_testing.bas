Attribute VB_Name = "Module1"
Sub Alphabetical_testing():

Dim ws As Worksheet

For Each ws In ThisWorkbook.Worksheets

Dim total_row As Long
Dim output_row As Integer
Dim stock_count As Long
Dim stock_amount As Double



ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"

output_row = 1
stock_count = 0
stock_amount = 0



    total_row = ws.Cells(Rows.Count, 1).End(xlUp).Row

 For i = 2 To total_row
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
       
           output_row = output_row + 1
        ws.Cells(output_row, 9).Value = ws.Cells(i, 1).Value
              
          stock_count = stock_count + 1
          opening_row = i - stock_count + 1
        ws.Cells(output_row, 10).Value = ws.Cells(i, 6).Value - ws.Cells(opening_row, 3).Value
          
          If ws.Cells(output_row, 10).Value < 0 Then
            ws.Cells(output_row, 10).Interior.Color = vbRed
          Else
            ws.Cells(output_row, 10).Interior.Color = vbGreen
            
          End If
            
        ws.Cells(output_row, 11).Value = (ws.Cells(i, 6).Value - ws.Cells(opening_row, 3).Value) / ws.Cells(opening_row, 3).Value
        ws.Cells(output_row, 11).NumberFormat = "0.00%"
            
           If ws.Cells(output_row, 11).Value < 0 Then
            ws.Cells(output_row, 11).Interior.Color = vbRed
           Else
            ws.Cells(output_row, 11).Interior.Color = vbGreen
           End If
    
    
          stock_amount = stock_amount + ws.Cells(i, 7).Value
        ws.Cells(output_row, 12).Value = stock_amount
        
            stock_count = 0
            stock_amount = 0
    Else
        
        stock_count = stock_count + 1

        
        stock_amount = stock_amount + ws.Cells(i, 7).Value
    End If
 
 Next i




'-------------------------------------------------------------------------------------
'Add Functionality to your script to return the stock with the "Greatest % increase",
'"Greatest % decrease", and "Greatest total volume",
'------------------------------------------------------------------------------------



Dim row_count As Integer


row_count = ws.Cells(Rows.Count, 9).End(xlUp).Row

ws.Cells(1, 16).Value = "Ticker"
ws.Cells(1, 17).Value = "Value"
ws.Cells(2, 15).Value = "Greatest % Increase"
ws.Cells(3, 15).Value = "Greatest % Decrease"
ws.Cells(4, 15).Value = "Greatest Total Volume"

 
  ws.Cells(2, 17).Value = WorksheetFunction.Max(ws.Range("K2:K" & row_count))
  ws.Cells(2, 17).NumberFormat = "0.00%"

  
  ws.Cells(3, 17).Value = WorksheetFunction.Min(ws.Range("K2:k" & row_count))
  ws.Cells(3, 17).NumberFormat = "0.00%"

  ws.Cells(4, 17).Value = WorksheetFunction.Max(ws.Range("L2:L" & row_count))
  ws.Cells(4, 17).NumberFormat = "##0.00E+0"
  
    For i = 2 To row_count
        If Cells(i, 11).Value = ws.Cells(2, 17).Value Then
          ws.Cells(2, 16).Value = ws.Cells(i, 9).Value

        ElseIf Cells(i, 11).Value = ws.Cells(3, 17).Value Then
           ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
         
        ElseIf Cells(i, 12).Value = ws.Cells(4, 17).Value Then
           ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
           
        End If
     
     Next i
  
 Next ws
     
End Sub


