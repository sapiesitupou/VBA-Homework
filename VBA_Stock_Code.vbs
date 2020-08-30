Sub stock_data()
 
    Dim MyWS As Worksheet
    Dim Summary_Header As Boolean
    Dim Command_Sheet As Boolean
    
    For Each MyWS In Worksheets
    
        Dim ticker_letter As String
        ticker_letter = " "
        
        Dim Final_Ticker As Double
        Final_Ticker = 0
        
       
        Dim open_price As Double
        open_price = 0
        Dim close_price As Double
        close_price = 0
        Dim ChangeInPrice As Double
        ChangeInPrice = 0
        Dim PercentageChange As Double
        PercentageChange = 0
       
        Dim MAX_ticker_letter As String
        MAX_ticker_letter = " "
        Dim MIN_ticker_letter As String
        MIN_ticker_letter = " "
        Dim MAX_PERCENT As Double
        MAX_PERCENT = 0
        Dim MIN_PERCENT As Double
        MIN_PERCENT = 0
        Dim MAX_VOLUME_TICKER As String
        MAX_VOLUME_TICKER = " "
        Dim MAX_VOLUME As Double
        MAX_VOLUME = 0
        '----------------------------------------------------------------
         
        Dim Summary_Table_Row As Long
        Summary_Table_Row = 2
        
        Dim Lastrow As Long
        Dim i As Long
        
        Lastrow = MyWS.Cells(Rows.Count, 1).End(xlUp).Row

        If Summary_Header Then
            MyWS.Range("I1").Value = "Ticker"
            MyWS.Range("J1").Value = "Yearly Change"
            MyWS.Range("K1").Value = "Percent Change"
            MyWS.Range("L1").Value = "Total Stock Volume"
            
            MyWS.Range("O2").Value = "Greatest % Increase"
            MyWS.Range("O3").Value = "Greatest % Decrease"
            MyWS.Range("O4").Value = "Greatest Total Volume"
            MyWS.Range("P1").Value = "Ticker"
            MyWS.Range("Q1").Value = "Value"
        Else
            Summary_Header = True
        End If
        
          open_price = MyWS.Cells(2, 3).Value
        
        For i = 2 To Lastrow
            If MyWS.Cells(i + 1, 1).Value <> MyWS.Cells(i, 1).Value Then
                ticker_letter = MyWS.Cells(i, 1).Value
                close_price = MyWS.Cells(i, 6).Value
                ChangeInPrice = close_price - open_price
                If open_price <> 0 Then
                    PercentageChange = (ChangeInPrice / open_price) * 100
                Else
            End If
                
                Final_Ticker = Final_Ticker + MyWS.Cells(i, 7).Value
              
                MyWS.Range("I" & Summary_Table_Row).Value = ticker_letter
                MyWS.Range("J" & Summary_Table_Row).Value = ChangeInPrice
                If (ChangeInPrice > 0) Then
                    MyWS.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                ElseIf (ChangeInPrice <= 0) Then
                    MyWS.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                End If
                
                MyWS.Range("K" & Summary_Table_Row).Value = (CStr(PercentageChange) & "%")
                MyWS.Range("L" & Summary_Table_Row).Value = Final_Ticker
                
               Summary_Table_Row = Summary_Table_Row + 1
                ChangeInPrice = 0
                close_price = 0
                open_price = MyWS.Cells(i + 1, 3).Value
              
                If (PercentageChange > MAX_PERCENT) Then
                    MAX_PERCENT = PercentageChange
                    MAX_ticker_letter = ticker_letter
                ElseIf (PercentageChange < MIN_PERCENT) Then
                    MIN_PERCENT = PercentageChange
                    MIN_ticker_letter = ticker_letter
                End If
                       
                If (Final_Ticker > MAX_VOLUME) Then
                    MAX_VOLUME = Final_Ticker
                    MAX_VOLUME_TICKER = ticker_letter
                End If
                
                PercentageChange = 0
                Final_Ticker = 0
                
            Else
                Final_Ticker = Final_Ticker + MyWS.Cells(i, 7).Value
            End If
           
      
        Next i

           If Command_Sheet Then
            
                MyWS.Range("Q2").Value = (CStr(MAX_PERCENT) & "%")
                MyWS.Range("Q3").Value = (CStr(MIN_PERCENT) & "%")
                MyWS.Range("P2").Value = MAX_ticker_letter
                MyWS.Range("P3").Value = MIN_ticker_letter
                MyWS.Range("Q4").Value = MAX_VOLUME
                MyWS.Range("P4").Value = MAX_VOLUME_TICKER
                
            Else
                Command_Sheet = True
            End If
        
     Next MyWS
     
End Sub


