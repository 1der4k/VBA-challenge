VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub TickerSymbol():

    'This website was a useful source for looping through worksheets script:http://excelhelphq.com/how-to-loop-through-worksheets-in-a-workbook-in-excel-vba/
    Dim ws As Worksheet
    Dim starting_ws As Worksheet
    Set starting_ws = ActiveSheet
    
    For Each ws In ThisWorkbook.Worksheets

'Label Columns
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Yearly Change"
        Cells(1, 11).Value = "Percent Change"
        Cells(1, 12).Value = "Total Stock Volume"
    
        
'Capture ticker symbol

        ws.Activate
                
        Dim LastRow As Long
        LastRow = ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Row
        
              
        Dim TickSym As String
        Dim SummaryRow As Integer
                
        SummaryRow = 2
        TickSym = ActiveSheet.Cells(SummaryRow, 1).Value
        
        Dim YearChange As Double
        Dim LastDayClose As Double
        Dim FirstDayOpen As Double
        Dim PercentYearChange As Double
        
        Dim StartRow As Integer
        FirstDayOpen = ActiveSheet.Cells(2, 3)
        
        Dim StockVolume As Double
        StockVolume = ActiveSheet.Cells(2, 7).Value
          
            
            For j = 2 To LastRow
                If ActiveSheet.Cells(j, 1).Value = ActiveSheet.Cells(j + 1, 1).Value Then
                    TickSym = ActiveSheet.Cells(j + 1, 1).Value
                    ActiveSheet.Cells(SummaryRow, 9) = TickSym
                    StockVolume = StockVolume + ActiveSheet.Cells(j + 1, 7).Value
                    
                ElseIf ActiveSheet.Cells(j, 1).Value <> ActiveSheet.Cells(j + 1, 1).Value Then
                    TickSym = ActiveSheet.Cells(j + 1, 1).Value
                    ActiveSheet.Cells(SummaryRow + 1, 9) = TickSym
                    LastDayClose = ActiveSheet.Cells(j, 6).Value
                    YearChange = LastDayClose - FirstDayOpen
                    ActiveSheet.Cells(SummaryRow, 10) = YearChange
                    If FirstDayOpen = 0 Then
                        PercentYearChange = 0
                    Else:
                        PercentYearChange = (YearChange / FirstDayOpen) * 100
                    End If
                    ActiveSheet.Cells(SummaryRow, 11) = PercentYearChange
                    ActiveSheet.Cells(SummaryRow, 12) = StockVolume
                    StockVolume = 0
                    FirstDayOpen = ActiveSheet.Cells(j + 1, 3).Value
                    SummaryRow = SummaryRow + 1
                    
                End If
            Next j
    Next
        
        
        
   

End Sub
