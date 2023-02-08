Attribute VB_Name = "Module1"
Sub STOCKDATA()

                            Dim TOT_VOL As Double: TOT_VOL = 0
                            Dim Row As Long
                            Dim Row_Count As Long
                            Dim Yearly_Change As Double: Yearly_Change = 0
                            Dim Percent_change As Double
                            Dim Stock_Row As Long: Stock_Row = 2
                            Dim SUMT_Row As Long: SUMT_Row = 0
               
                            
   For Each ws In Worksheets
    
        'Labels
                            ws.Range("I1").Value = "Ticker"
                            ws.Range("J1").Value = "Yearly Change"
                            ws.Range("K1").Value = "Percent Change"
                            ws.Range("L1").Value = "Total Stock Volume"
                            ws.Range("P1").Value = "Ticker"
                            ws.Range("Q1").Value = "Value"
                            ws.Range("O2").Value = "Greatest % Increase"
                            ws.Range("O3").Value = "Greatest % Decrease"
                            ws.Range("O4").Value = "Greatest Total Volume"
                          
                            Row_Count = ws.Cells(Rows.Count, "A").End(xlUp).Row
              For Row = 2 To Row_Count
                            
                        If ws.Cells(Row + 1, 1).Value <> ws.Cells(Row, 1).Value Then
                                            TOT_VOL = TOT_VOL + Cells(Row, 7).Value
                            
                                If TOL_VOL = 0 Then
                                    
                                    ws.Range("I" & 2 + SUMT_Row).Value = ws.Cells(Row, 1).Value
                                    ws.Range("J" & 2 + SUMT_Row).Value = 0
                                    ws.Range("K" & 2 + SUMT_Row).Value = 0 & "%"
                                    ws.Range("L" & 2 + SUMT_Row).Value = 0
                                
                                    If ws.Cells(Stock_Row, 3).Value = 0 Then
                                         For findValue = Stock_Row To Row
                                                 If ws.Cells(findValue, 3).Value <> 0 Then
                                                      Stock_Row = findValue
                                        
                                    End If
                            Next findValue
                        End If
                            
                                    Yearly_Change = (ws.Cells(Row, 6).Value - ws.Cells(Row, 3).Value)
                                    Percent_change = Yearly_Change / ws.Cells(Stock_Row, 3).Value
                
                                ws.Range("I" & 2 + SUMT_Row).Value = ws.Cells(Row, 1).Value
                                ws.Range("J" & 2 + SUMT_Row).Value = Yearly_Change
                                ws.Range("J" & 2 + SUMT_Row).NumberFormat = "00.00"
                                ws.Range("K" & 2 + SUMT_Row).Value = Percent_change
                                ws.Range("K" & 2 + SUMT_Row).NumberFormat = "00.00%"
                                ws.Range("L" & 2 + SUMT_Row).Value = TOT_VOL
                                ws.Range("L" & 2 + SUMT_Row).NumberFormat = "#,## "
                            
                     If Yearly_Change > 0 Then
                                    ws.Range("J" & 2 + SUMT_Row).Interior.ColorIndex = 4
                                    
                                    ElseIf Yearly_Change < 0 Then
                                ws.Range("J" & 2 + SUMT_Row).Interior.ColorIndex = 3
                                  
                                Else
                                    ws.Range("J" & 2 + SUMT_Row).Interior.ColorIndex = 0
                                    End If
                                    
                                 End If
                                    TOL_VOL = 0
                                    Yearly_Change = 0
                                    SUMT_Row = SUMT_Row + 1
                                 Else
                                    TOT_VOL = TOT_VOL + ws.Cells(Row, 7).Value
                                End If
                Next Row
                
                    ws.Range("Q2").Value = "%" & WorksheetFunction.Max(ws.Range("K2:K" & Row_Count)) * 100
                    ws.Range("Q3").Value = "%" & WorksheetFunction.Min(ws.Range("K2:K" & Row_Count)) * 100
                    ws.Range("Q4").Value = "%" & WorksheetFunction.Max(ws.Range("L2:L" & Row_Count)) * 100
                    ws.Range("Q4").NumberFormat = "#,##"
                                
                                 inumber = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K2:K" & Row_Count)), ws.Range("K2:K" & Row_Count), 0)
                                ws.Range("P2").Value = ws.Cells(inumber + 1, 9)
   
                                 dnumber = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K2:K" & Row_Count)), ws.Range("K2:K" & Row_Count), 0)
                                ws.Range("P3").Value = ws.Cells(dnumber + 1, 9)
                                
                                vnumber = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("L2:L" & Row_Count)), ws.Range("L2:L" & Row_Count), 0)
                                ws.Range("P3").Value = ws.Cells(vnumber + 1, 9)
   
   'autofit
   
                                ws.Columns("A:Q").AutoFit
                                            
        
    
    Next ws

End Sub

