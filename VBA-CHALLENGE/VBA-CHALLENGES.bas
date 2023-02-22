Attribute VB_Name = "Module1"
Sub STOCKDATA():

'Declare
Dim TICKER_NAME As String
Dim TICKER_NUM As Integer
Dim LastRow As Long
Dim OP As Double
Dim CP As Double
Dim YEARLY_CHANGE As Double
Dim PERCENT_CHANGE As Double
Dim TOTAL_S_VOL  As Double
Dim GREATEST_PERCENT_INC As Double
Dim GREATEST_PERCENT_INC_T As String
Dim GREATEST_PERCENT_DEC As Double
Dim GREATEST_PERCENT_DEC_T As String
Dim GREATEST_STOCK_VOLUME As Double
Dim GREATEST_STOCK_VOLUME_T As String

' loop in the workbook
For Each ws In Worksheets
    ws.Activate

    'Last Row
    LastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row

    'Header
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    
    'Varibles
    TICKER_NUM = 0
    TICKER_NAME = ""
    YEARLY_CHANGE = 0
    OP = 0
    PERCENT_CHANGE = 0
    TOTAL_S_VOL = 0
    
    'ticker loop
    For i = 2 To LastRow
        TICKER_NAME = Cells(i, 1).Value
        
        If OP = 0 Then
            OP = Cells(i, 3).Value
        End If
        
        TOTAL_S_VOL = TOTAL_S_VOL + Cells(i, 7).Value
        
        If Cells(i + 1, 1).Value <> TICKER_NAME Then
            TICKER_NUM = TICKER_NUM + 1
            Cells(TICKER_NUM + 1, 9) = TICKER_NAME
         
            CP = Cells(i, 6)
            
            YEARLY_CHANGE = CP - OP
            
            Cells(TICKER_NUM + 1, 10).Value = YEARLY_CHANGE
            
           'Color
            If YEARLY_CHANGE > 0 Then
                Cells(TICKER_NUM + 1, 10).Interior.ColorIndex = 4
            ElseIf YEARLY_CHANGE < 0 Then
                Cells(TICKER_NUM + 1, 10).Interior.ColorIndex = 3
            Else
                Cells(TICKER_NUM + 1, 10).Interior.ColorIndex = 6
            End If
            
            
            'percent change
            If OP = 0 Then
                PERCENT_CHANGE = 0
            Else
                PERCENT_CHANGE = (YEARLY_CHANGE / OP)
            End If
            
            Cells(TICKER_NUM + 1, 11).Value = Format(PERCENT_CHANGE, "Percent")
            OP = 0
            Cells(TICKER_NUM + 1, 12).Value = TOTAL_S_VOL
            TOTAL_S_VOL = 0
        End If
        
    Next i
    
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
    
    
    LastRow = ws.Cells(Rows.Count, "I").End(xlUp).Row
    
    GREATEST_PERCENT_INC = Cells(2, 11).Value
    GREATEST_PERCENT_INC_T = Cells(2, 9).Value
    GREATEST_PERCENT_DEC = Cells(2, 11).Value
    GREATEST_PERCENT_DEC_T = Cells(2, 9).Value
    GREATEST_STOCK_VOLUME = Cells(2, 12).Value
    GREATEST_stock_volume_ticker = Cells(2, 9).Value
    
    
   
    For i = 2 To LastRow
        If Cells(i, 11).Value > GREATEST_PERCENT_INC Then
            GREATEST_PERCENT_INC = Cells(i, 11).Value
            GREATEST_PERCENT_INC_T = Cells(i, 9).Value
        End If
        
       
        If Cells(i, 11).Value < GREATEST_PERCENT_DEC Then
            GREATEST_PERCENT_DEC = Cells(i, 11).Value
            GREATEST_PERCENT_DEC_T = Cells(i, 9).Value
        End If
        
      
        If Cells(i, 12).Value > GREATEST_STOCK_VOLUME Then
            GREATEST_STOCK_VOLUME = Cells(i, 12).Value
            GREATEST_STOCK_VOLUME_T = Cells(i, 9).Value
        End If
        
    Next i
    
    
    Range("P2").Value = Format(GREATEST_PERCENT_INC_T, "Percent")
    Range("Q2").Value = Format(GREATEST_PERCENT_INC, "Percent")
    Range("P3").Value = Format(GREATEST_PERCENT_DEC_T, "Percent")
    Range("Q3").Value = Format(GREATEST_PERCENT_DEC, "Percent")
    Range("P4").Value = GREATEST_STOCK_VOLUME_T
    Range("Q4").Value = GREATEST_STOCK_VOLUME
    
Next ws


End Sub
