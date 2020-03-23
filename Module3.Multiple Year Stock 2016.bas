Attribute VB_Name = "Module3"
Sub Stock_Market_2016()

Dim Ticker_Name As String

Dim Total_Stock_Volume As Double
Total_Stock_Volume = 0

Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2

Dim Yearly_Change As Double
        
Dim change As Double
    change = 0
    
Dim Percent_Change As Double

    
Range("J1").Value = "Yearly_Change"

    For i = 2 To 797711
        
        
        If Cells(i + 1, 1).Value <> Cells(i, 1) Then
        
            Ticker_Name = Cells(i, 1).Value
            
            Total_Stock_Volume = Total_Stock_Volume + Cells(i, 12).Value
            
            Range("I" & Summary_Table_Row).Value = Ticker_Name
                        
            Range("L" & Summary_Table_Row).Value = Total_Stock_Volume

      
            Summary_Table_Row = Summary_Table_Row + 1
      
      
      Total_Stock_Volume = 0

        Else
                change = (Cells(i, 6) - Cells(i + 1, 3))
                Range("J" & Summary_Table_Row).Value = change
        
               percentChange = Round((change / Cells(i + 1, 3) * 100), 2)
                     
               Range("K" & Summary_Table_Row).Value = "%" & percentChange
                              
                     
            Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value

    End If

  Next i

End Sub


