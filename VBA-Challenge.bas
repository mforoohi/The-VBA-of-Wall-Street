Attribute VB_Name = "Module2"
Sub StockChangeCalculater():

Dim ws As Worksheet

For Each ws In ThisWorkbook.Sheets

    Dim ticker_symbol As String
       ticker_symbol = " "

    Dim yearly_change As Double

    Dim percent_change As Double

    Dim total_stock_volume As Double
        'total_stock_volume = 0

    Dim year_open As Double
       'year_open = 0

    Dim year_close As Double
       'year_close = 0

    Dim summary_table_row As Integer
       summary_table_row = 2
       
    Dim Lastrow As Long
    Lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row




        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"

    For i = 2 To Lastrow

         If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
               ticker_symbol = ws.Cells(i, 1).Value
               year_open = ws.Cells(i, 3).Value
               year_close = ws.Cells(i, 6).Value
        
                yearly_change = (year_close - year_open)
        
            If year_close <> 0 Then
              percent_change = (yearly_change / year_close) * 100
            End If
       

               total_stock_volume = total_stock_volume + ws.Cells(i, 7).Value
    
              ws.Range("I" & summary_table_row).Value = ticker_symbol
              ws.Range("j" & summary_table_row).Value = yearly_change
              ws.Range("K" & summary_table_row).Value = percent_change
              ws.Range("L" & summary_table_row).Value = total_stock_volume
    
              summary_table_row = summary_table_row + 1
    
              tiker_symbol = ""
              yearly_change = 0
              percent_change = 0
              total_stock_volume = 0
         
    
    Else
        
        total_stock_volume = total_stock_volume + Cells(i, 7).Value
        
    End If
    
Next i

       For i = 2 To ws.UsedRange.Rows.Count
    
          
             If ws.Cells(i, 10).Value >= 0 Then

                    ws.Cells(i, 10).Interior.ColorIndex = 4
    
             Else
                   ws.Cells(i, 10).Interior.ColorIndex = 3
            End If
                    
      Next i
       
Next ws

End Sub



