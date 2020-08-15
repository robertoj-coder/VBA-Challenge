Sub Multiyear_stock_data()
  
  ' LOOP THROUGH ALL SHEETS
'------------------------------
       
    For Each ws In Worksheets
    
' Created a Variable to store ticker, yearly price, yearly percentage , total volume


    Dim Ticker As String
    Dim OpeningPrice As Double
    Dim ClosingPrice As Double
    Dim YearlyChange As Double
    Dim PercentChange As Variant
    Dim total_stock_volume As Variant

    OpeningPrice = 0
    ClosingPrice = 0
    YearlyChange = 0
    PercentChange = 0
    total_stock_volume = 0
    
    
    ' Keep track of the location for each ticker in the table row
    Dim Table_Row As Long
    Table_Row = 2
    
    'Determine Final Row
    
    Lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    'Set initial value of OpeningPrice for the first Ticker
    
    OpeningPrice = ws.Cells(2, 3).Value
    
    'MsgBox for first opening price, allow the system to not freeze
    
    MsgBox (OpeningPrice)
    
    'Loop through all tickers
    For i = 2 To Lastrow
    
' Check if we are still within the same Ticker
    
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
        Ticker = ws.Cells(i, 1).Value
        
        'Determine Yearly Change
            
        ClosingPrice = ws.Cells(i, 6).Value
        
        YearlyChange = ClosingPrice - OpeningPrice
        
        ' Determine Percentage change and check division by zero avoid error overflow
        
        If OpeningPrice = 0 Then
        
        PercentChange = 0
        Else
       
        
        PercentChange = (YearlyChange / OpeningPrice)
        
        End If
        
        'Determine Total stock Volume
        
        total_stock_volume = total_stock_volume + ws.Cells(i, 7).Value

    ' Print The Ticker , yearly change , percentage change and volume to Table Row/formatting to percentage
        ws.Range("J" & Table_Row).Value = Ticker
        ws.Range("K" & Table_Row).Value = YearlyChange
        ws.Range("L" & Table_Row).Value = PercentChange
        ws.Range("L" & Table_Row).NumberFormat = "0.00%"
        ws.Range("M" & Table_Row).Value = total_stock_volume
        
    ' Fill "Yearly Change" with green to  positive change and red to negative change
        If YearlyChange > 0 Then
        ws.Range("K" & Table_Row).Interior.ColorIndex = 4
        ElseIf YearlyChange <= 0 Then
        ws.Range("K" & Table_Row).Interior.ColorIndex = 3
        End If

       
    
    ' Set Titles for the Table_Row for each worksheet
    
        ws.Cells(1, 10).Value = "Ticker"
        ws.Cells(1, 11).Value = "Yealy Change"
        ws.Cells(1, 12).Value = "PercentChange"
        ws.Cells(1, 13).Value = " Total Stock Volume"
        
    'Add 1 to the table row
        
        Table_Row = Table_Row + 1
    
    'Reset to 0
        
        YearlyChange = 0
            
        ClosingPrice = 0
        
    ' Capture new opening price
        
        OpeningPrice = ws.Cells(i + 1, 3).Value
    
    'reset total volume to 0
        
        total_stock_volume = 0
  
        'Add to the Volume total
      Else
        
        total_stock_volume = total_stock_volume + ws.Cells(i, 7).Value
        
        
     End If
    
      Next i
      
        'Second Part Challenge..............................................................
      
      Lastrow = ws.Cells(Rows.Count, 12).End(xlUp).Row
      
      'Start Loop For last challenge
      
      For i = 2 To Lastrow
      
      'Determine the Greatest increase and decrease
   
        If ws.Cells(i, 12) > ws.Cells(2, 17) Then
        
            ws.Cells(2, 17) = ws.Cells(i, 12)
            ws.Cells(2, 16) = ws.Cells(i, 10)
            
        ElseIf ws.Cells(i, 12) < ws.Cells(3, 17) Then
         ws.Cells(3, 17) = ws.Cells(i, 12)
         ws.Cells(3, 16) = ws.Cells(i, 10)
         
        'Format print to percentage
        
        ws.Range("Q2").NumberFormat = "0.00%"
        ws.Range("Q3").NumberFormat = "0.00%"
        
        End If
        
        'Determine the greates total volume
        
        If ws.Cells(i, 13) > ws.Cells(4, 17) Then
        
        ws.Cells(4, 17) = ws.Cells(i, 13)
        ws.Cells(4, 16) = ws.Cells(i, 10)
    
        
            
       
        End If
        
       
       Next i
       
       'Add headers to the second part Challenge
      
        ws.Cells(2, 15).Value = "Greatest % increase"
        ws.Cells(3, 15).Value = "Greatest % decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"

     
        

        Next ws
   

End Sub
