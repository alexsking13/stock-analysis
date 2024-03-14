Attribute VB_Name = "Module1"
Sub RunSubAcrossWorksheets()
    Dim ws As Worksheet
    
        For Each ws In ThisWorkbook.Sheets
            StockAnalysis ws
        Next ws
End Sub
Sub StockAnalysis(ws As Worksheet)

    Dim rowNumber As Long
    Dim lastRowNum As Long
    Dim tickerNum As Integer
    Dim tickerName As String
    Dim openingPrice As Double
    Dim closingPrice As Double
    Dim percentChange As Double
    Dim priceDifference As Double
    Dim stockVolume As Double
    Dim greatestpercentinc As Double
    Dim greatestpercentdec As Double
    Dim greatesttotalvol As Double
    Dim tickgreatestpercinc As String
    Dim tickgreatestpercdec As String
    Dim tickggreatestotalvol As String

   
   
        ws.Cells(1, "J") = "Ticker"
        ws.Cells(1, "K") = "Price Difference"
        ws.Cells(1, "L") = "Percent Change"
        ws.Cells(1, "M") = "Total Stock Volume"
        ws.Cells(1, "Q") = "Ticker"
        ws.Cells(1, "R") = "Value"
        ws.Cells(2, "P") = "Greatest % Increase"
        ws.Cells(3, "P") = "Greatest % Decrease"
        ws.Cells(4, "P") = "Greatest Total Volume"
        


    lastRowNum = ws.Cells(Rows.Count, "A").End(xlUp).Row
    greatestpercentinc = 0
    greatestpercentdec = 0
    greatesttotalvol = 0
    

    tickerNum = 2
    For rowNumber = 2 To lastRowNum
        If ws.Cells(rowNumber, 1) <> tickerName Then
            openingPrice = ws.Cells(rowNumber, 3)
            stockVolume = 0
            tickerName = ws.Cells(rowNumber, 1)
        End If
        ws.Cells(tickerNum, 10) = tickerName
        
        closingPrice = ws.Cells(rowNumber, 6)
        stockVolume = stockVolume + ws.Cells(rowNumber, 7)
        
        
        If ws.Cells(rowNumber, 1) <> ws.Cells(rowNumber + 1, 1) Then

            percentChange = (closingPrice - openingPrice) / openingPrice
            priceDifference = closingPrice - openingPrice
            If priceDifference > 0 Then
            ws.Cells(tickerNum, 11).Interior.Color = RGB(0, 255, 0)
            ElseIf priceDifference < 0 Then
            ws.Cells(tickerNum, 11).Interior.Color = RGB(255, 0, 0)
            End If
            ws.Cells(tickerNum, 11) = priceDifference
            ws.Cells(tickerNum, 12) = percentChange
            ws.Cells(tickerNum, 13) = stockVolume
            
            If percentChange > greatestpercentinc Then
            greatestpercentinc = percentChange
            tickgreatestpercinc = tickerName
            End If
            ws.Cells(2, "R") = greatestpercentinc
            ws.Cells(2, "Q") = tickgreatestpercinc
            ws.Range("R2").NumberFormat = "0.00%"
            
            If percentChange < greatestpercentdec Then
            greatestpercentdec = percentChange
            tickgreatestpercdec = tickerName
            End If
            ws.Cells(3, "R") = greatestpercentdec
            ws.Cells(3, "Q") = tickgreatestpercdec
            ws.Range("R3").NumberFormat = "0.00%"
            
            If stockVolume > greatesttotalvol Then
            greatesttotalvol = stockVolume
            tickggreatestotalvol = tickerName
            End If
            ws.Cells(4, "R") = greatesttotalvol
            ws.Cells(4, "Q") = tickggreatestotalvol
            
            ws.Range("L2:L" & lastRowNum).NumberFormat = "0.00%"
            
            tickerNum = tickerNum + 1
            
            
        End If
    Next rowNumber
End Sub
