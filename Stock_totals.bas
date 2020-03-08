Attribute VB_Name = "Module1"
Sub TickerandVol():

' Set an initial variable for holding the ticker name
Dim Ticker_Type As String
Dim number_tickers As Integer
Dim lastRowState As Long
Dim opening_price As Double
Dim closing_price As Double
Dim yearly_change As Double
Dim percent_change As Double
Dim total_stock_volume As Double

' loop each worksheet in the workbook
For Each ws In Worksheets
  
'Make the worksheet active.
ws.Activate

' whats last row of each worksheet
    lastRowState = ws.Cells(Rows.Count, "A").End(xlUp).Row
    
    number_tickers = 0
    ticker = ""
    yearly_change = 0
    opening_price = 0
    percent_change = 0
    total_stock_volume = 0
    
  ' Loop through all tickers
 
  For i = 2 To lastRowState

         ticker = Cells(i, 1).Value
         
        ' Set the ticker name
         Ticker_Type = Cells(i, 1).Value

         ' Add to the ticker name Total
           Ticker_Type = Ticker_Type
                    yearly_change = Cells(i, 3).Value - Cells(i, 6).Value
      
          ' Add to the vol Total
            Vol_Total = Vol_Total + Cells(i, 7).Value
                    Vol_Total = 0
        
        ' Get the start of the year price
            If opening_price = 0 Then
            opening_price = Cells(i, 3).Value
            
        End If
        
        ' Add up the total stock volume values for a ticker.
            total_stock_volume = total_stock_volume + Cells(i, 7).Value
        
        ' Run this if we get to a different ticker in the list.
             If Cells(i + 1, 1).Value <> ticker Then
             
                number_tickers = number_tickers + 1
                Cells(number_tickers + 1, 9) = ticker
            
            ' Get the end of the year closing price for ticker
                closing_price = Cells(i, 6)
            
            ' Get yearly change value
                yearly_change = closing_price - opening_price
            
            ' yearly change value
                 Cells(number_tickers + 1, 10).Value = yearly_change
            
            ' shade cell green.
            
                 If yearly_change > 0 Then
                Cells(number_tickers + 1, 10).Interior.ColorIndex = 4
                
            ' If yearly change value is less than 0, shade cell red.
                ElseIf yearly_change < 0 Then
                    Cells(number_tickers + 1, 10).Interior.ColorIndex = 3
                    
            ' If yearly change value is 0, shade cell yellow.
            
            
            Else
            
                Cells(number_tickers + 1, 10).Interior.ColorIndex = 6
                
                
            End If
            
        
            ' find percent change value for ticker.
                 If opening_price = 0 Then
                percent_change = 0
                
            Else
            
                percent_change = (yearly_change / opening_price)
                
            End If
            
            
            ' Format the percent_change value as a percent.
            Cells(number_tickers + 1, 11).Value = Format(percent_change, "Percent")
            
            ' Set opening price back to 0
            opening_price = 0
            
            ' total stock volume
            Cells(number_tickers + 1, 12).Value = total_stock_volume
            
            ' set to zero
            total_stock_volume = 0
            
        End If
        
    Next i
  
            If opening_price = 0 Then
                percent_change = 0
                
Else
                percent_change = (yearly_change / opening_price)
                
            End If
            
  Next ws
            

End Sub

