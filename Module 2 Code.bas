Attribute VB_Name = "Module1"
Option Explicit


Sub Challenge2()


Dim TickerSymbol As String
Dim YearlyChange As Double
Dim PercentChange As Double

Dim OpeningPrice As Double

Dim ClosingPrice As Double

Dim TSrow As Long
Dim LastRow As Long
Dim LastColumn As Integer
Dim i As Long
Dim TotalVolume As Double



TotalVolume = 0



LastRow = Cells(Rows.Count, 1).End(xlUp).Row
LastColumn = Cells(Columns.Count).End(xlToLeft).Column
TSrow = 2
OpeningPrice = Cells(2, 3).Value

    
'add headers columns I-L & P-Q
Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly Change"
Range("K1").Value = "Percent Change"
Range("L1").Value = "Total Stock Volume"
        
          
            
             
            For i = 2 To LastRow
                                            
                        'Compare current cell ith the next to see if they are different
                        If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
                       
                        
                        'Add value in current cell to column I (the next non-empty row
                                
                        TickerSymbol = Cells(i, 1).Value
                        
                        Cells(TSrow, 9).Value = TickerSymbol
                        
                               
                        'note closing price for i stock
                        ClosingPrice = Cells(i, 6).Value
                        
                        'note the running total volume
                        TotalVolume = Cells(i, 7).Value + TotalVolume
                        'add total volume to colum L
                        Cells((TSrow), 12).Value = TotalVolume
                        
                        
                        'calculate the yearly change
                        YearlyChange = (ClosingPrice - OpeningPrice)
                        
                        Cells(TSrow, 10).Value = YearlyChange
                        
                        'Add Percentage Change -'Percent Change = (New Value – Old Value) / Old Value* 100
                        PercentChange = (YearlyChange / OpeningPrice)
                                    
                       
                        OpeningPrice = Cells(i + 1, 3).Value
                        
                        
                        Range("K:K").NumberFormat = "0.00%"
                        Cells(TSrow, 11).Value = PercentChange
                        
                                                                       
                        
                        'Reset Total Volume
                        TotalVolume = 0
                        TSrow = TSrow + 1
                        
                        
                         
                                           
                        Else
                        
                        TotalVolume = Cells(i, 7).Value + TotalVolume
                                   
                        End If
                        
                        'Add color
                        If Cells(i, 10).Value > 0 Then
                        Cells(i, 10).Interior.ColorIndex = 4
                        
                        ElseIf Cells(i, 10).Value < 0 Then
                        Cells(i, 10).Interior.ColorIndex = 3
                        
                        Else
                        Cells(i, 10).Interior.ColorIndex = 2
                        
                        
                        End If

            
            Next i

            

 
End Sub

