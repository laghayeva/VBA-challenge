Attribute VB_Name = "Module1"
Sub StockAnalysis()

    'Looping through all worksheets
    
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets

    ' Adding Headers
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Quarterly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"

    ' Defining Variables
    Dim Ticker As String
    Dim Quarterly_Change As Double
    Dim Percent_Change As Double
    Dim Total_Stock_Volume As Double
    Dim Summary_Row As Integer
    Dim Last_Row As Long
    Dim Opening_Price As Double
    Dim Closing_Price As Double

    ' Initializing Variables
    Quarterly_Change = 0
    Percent_Change = 0
    Summary_Row = 2
    Total_Stock_Volume = 0
    Last_Row = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Opening_Price = 0
    Closing_Price = 0

'Looping Through

    For i = 2 To Last_Row


       If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
            Opening_Price = ws.Cells(i, 3).Value
       End If


        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            Ticker = ws.Cells(i, 1).Value
            Closing_Price = ws.Cells(i, 6).Value
            
        
            Quarterly_Change = Closing_Price - Opening_Price
            ws.Range("J" & Summary_Row).Value = Quarterly_Change

         
            If Opening_Price <> 0 Then
            Percent_Change = (Quarterly_Change / Opening_Price)
            Else
                Percent_Change = 0
            End If
            ws.Range("K" & Summary_Row).Value = Percent_Change
            ws.Range("K" & Summary_Row).NumberFormat = "0.00%"


            ws.Range("I" & Summary_Row).Value = Ticker
            Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
            ws.Range("L" & Summary_Row).Value = Total_Stock_Volume

            Summary_Row = Summary_Row + 1
            Total_Stock_Volume = 0
            Opening_Price = 0
            Closing_Price = 0
        Else
           
            Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
        End If

    Next i
    
'Looping for the final summary

 Dim Greatest_Percent_Increase As Double
    Greatest_Percent_Increase = -1000000

    For i = 2 To Last_Row
        If ws.Cells(i, 11).Value > Greatest_Percent_Increase Then
            Greatest_Percent_Increase = ws.Cells(i, 11).Value
            ws.Range("P2").Value = ws.Cells(i, 9).Value
            ws.Range("Q2").Value = Greatest_Percent_Increase
            ws.Range("Q2").NumberFormat = "0.00%"
        End If
    Next i

 
    Dim Greatest_Percent_Decrease As Double
    Greatest_Percent_Decrease = 1000000

    For i = 2 To Last_Row
        If ws.Cells(i, 11).Value < Greatest_Percent_Decrease Then
            Greatest_Percent_Decrease = ws.Cells(i, 11).Value
           ws.Range("P3").Value = ws.Cells(i, 9).Value
            ws.Range("Q3").Value = Greatest_Percent_Decrease
            ws.Range("Q3").NumberFormat = "0.00%"
        End If
    Next i

  
    Dim Greatest_Total_Volume As Double
    Greatest_Total_Volume = 0

    For i = 2 To Last_Row
        If ws.Cells(i, 12).Value > Greatest_Total_Volume Then
            Greatest_Total_Volume = ws.Cells(i, 12).Value
            ws.Range("P4").Value = ws.Cells(i, 9).Value
            ws.Range("Q4").Value = Greatest_Total_Volume
        End If
    Next i

For i = 2 To Last_Row

    If ws.Cells(i, 10).Value > 0 Then
        ws.Cells(i, 10).Interior.ColorIndex = 4
        ws.Cells(i, 11).Interior.ColorIndex = 4
    
    Else
        ws.Cells(i, 10).Interior.ColorIndex = 3
        ws.Cells(i, 11).Interior.ColorIndex = 3
    
    End If
Next i
    
    

    ws.Cells.Columns.AutoFit
     
Next ws

End Sub

