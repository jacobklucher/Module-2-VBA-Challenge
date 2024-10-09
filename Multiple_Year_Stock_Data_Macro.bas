Attribute VB_Name = "Module1"
Option Explicit

Sub Multiple_Year_Stock_Data_Macro()

    'Variables
    Dim ws As Worksheet

    Dim i As Double
    Dim j As Double
    
    Dim TickerName As String
    Dim QuarterChangeOpen As Double
    Dim QuarterChangeClose As Double
    Dim VolumeTotal As Double
    
    Dim LastRow As Double
    Dim SummaryTable As Integer
    
    
    Dim MaxValue As Double
    Dim MinValue As Double
    Dim MaxVolume As Double
    
    Dim MaxStock As String
    Dim MinStock As String
    Dim MaxVolumeStock As String
    
    Dim StockRange As Range
    Dim VolumeRange As Range
    
    
    'Loop through each worksheet in workbook
    For Each ws In ThisWorkbook.Worksheets
    
        ws.Activate
        
        'Set variables to default
        QuarterChangeOpen = 0
        QuarterChangeClose = 0
        VolumeTotal = 0
        
        MaxStock = ""
        MinStock = ""
        MaxVolumeStock = ""
        
        SummaryTable = 2
    
        'Find last row
        LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    
        'Add headers to current worksheet
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Quarterly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        
    
        'Loop through rows to calculate stock data
        For i = 2 To LastRow
        
            If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
            
                VolumeTotal = VolumeTotal + ws.Cells(i, 7).Value
                
                QuarterChangeOpen = ws.Cells(i, 3).Value
        
            ElseIf ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
                TickerName = ws.Cells(i, 1).Value
                
                VolumeTotal = VolumeTotal + ws.Cells(i, 7).Value
                
                
                QuarterChangeClose = ws.Cells(i, 6).Value
                
                'Output data to summary table
                ws.Range("I" & SummaryTable).Value = TickerName
                
                ws.Range("J" & SummaryTable).Value = QuarterChangeClose - QuarterChangeOpen
                
                'Format color change
                If (QuarterChangeClose - QuarterChangeOpen) > 0 Then
                
                    ws.Range("J" & SummaryTable).Interior.ColorIndex = 4 'Green
                    
                ElseIf (QuarterChangeClose - QuarterChangeOpen) < 0 Then
                    
                    ws.Range("J" & SummaryTable).Interior.ColorIndex = 3 'Red
                    
                End If
                
                
                ws.Range("K" & SummaryTable).Value = (QuarterChangeClose - QuarterChangeOpen) / QuarterChangeOpen
                
                ws.Range("K" & SummaryTable).NumberFormat = "0.00%"
                
                ws.Range("L" & SummaryTable).Value = VolumeTotal
                
                'Reset for next ticker
                SummaryTable = SummaryTable + 1
                
                VolumeTotal = 0
                
            Else
            
                VolumeTotal = VolumeTotal + ws.Cells(i, 7).Value
                
            End If
        
        Next i
    
        
        'Set stock range and volume range to the summary table
        Set StockRange = ws.Range("K2:K" & SummaryTable - 1)
        Set VolumeRange = ws.Range("L2:L" & SummaryTable - 1)
        
        'Find max and min values for percent change
        MaxValue = WorksheetFunction.Max(StockRange)
        MinValue = WorksheetFunction.Min(StockRange)
        MaxVolume = WorksheetFunction.Max(VolumeRange)
        
        
        'Find ticker associated with max/min values
        For j = 2 To SummaryTable - 1
            
            If ws.Cells(j, 11).Value = MaxValue Then
            
                MaxStock = ws.Cells(j, 9).Value
                
            End If
            
            If ws.Cells(j, 11).Value = MinValue Then
            
                MinStock = ws.Cells(j, 9).Value
                
            End If
            
            If ws.Cells(j, 12).Value = MaxVolume Then
            
                MaxVolumeStock = ws.Cells(j, 9).Value
            
            End If
            
        Next j
            
        'Output values to summary table
        ws.Cells(2, 17).Value = MaxValue
        ws.Cells(2, 16).Value = MaxStock
        
        ws.Cells(3, 17).Value = MinValue
        ws.Cells(3, 16).Value = MinStock
        
        ws.Cells(4, 17).Value = MaxVolume
        ws.Cells(4, 16).Value = MaxVolumeStock
        
        ws.Range("Q2:Q3").NumberFormat = "0.00%"
        
        ws.Range("Q4").NumberFormat = "0.00E+0"
    
    
    Next ws

End Sub
