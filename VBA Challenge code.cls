VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub stock()

' Loop through each worksheet
  For Each ws In Worksheets

    ' Set an initial variable for holding the ticker symbol
    Dim Ticker_Symbol As String
    Dim WorksheetNams As String
   
        ' Grabbed the WorksheetName
        WorksheetName = ws.Name

        ' Set an initial variable for holding the price change
        Dim Opening_Price As Double
        Opening_Price = 0
        Dim Closing_Price As Double
        Closing_Price = 0
        Dim Price_Change As Double
        Price_Change = 0
        Dim Percentage_Change As Double
        Percentage_Change = 0
        Dim Total_Vol As Double
        Total_Volume = 0
               
        ' Find the last row and last column in the sheet that contains data
          LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
          LastCol = ws.Cells(1, Columns.Count).End(xlToLeft).Column

              ' Keep track of the price change for each ticker symbol in the summary table
                Dim Summary_Table_Row As Integer
                Summary_Table_Row = 2
                ws.Cells(1, LastCol + 2).Value = "Year"
                ws.Cells(1, LastCol + 3).Value = "Ticker Symbol"
                ws.Cells(1, LastCol + 4).Value = "Price Difference"
                ws.Cells(1, LastCol + 5).Value = "Percentage Change"
                ws.Cells(1, LastCol + 6).Value = "Total Stock Volume"
               
                Opening_Price = ws.Cells(2, 3).Value

                 ' Loop through all ticker prices
                 For i = 2 To LastRow ' LastRow
                 
                    ' Check if we are still within the same ticker symbol
                    ' if it is not...
                        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                            ' Print the Worksheet name in hte Summary Table Column I
                                ws.Cells(Summary_Table_Row, LastCol + 2).Value = WorksheetName
                            ' Set the ticker symbol
                                Ticker_Symbol = ws.Cells(i, 1).Value
                               
                                ' Find the closing price
                                Closing_Price = ws.Cells(i, 6).Value
                                                                                 
                               
                                ' Difference between the opening and closing price
                                Price_Change = Closing_Price - Opening_Price
                                 ' Percentage change by end of the year
                                 If Opening_Price <> 0 Then
                                Percentage_Change = Price_Change / Opening_Price
                               
                                End If
                               
                                ' Total volume of the stock
                                Total_Volume = Total_Volume + ws.Cells(i, 7).Value
                                ' Print the ticker symbol and price change in the Summary Table column J & K
                                ws.Cells(Summary_Table_Row, LastCol + 3).Value = Ticker_Symbol
                                ' Print the Price Change to the Summary Table
                                ws.Cells(Summary_Table_Row, LastCol + 4).Value = Price_Change
                               
                                'Format the cell simulaneosuly to 2 decimal points and color
                                ws.Cells(Summary_Table_Row, LastCol + 5).Value = FormatPercent(Percentage_Change, 3)
                                    If ws.Cells(Summary_Table_Row, LastCol + 4).Value > 0 Then
                                    ws.Cells(Summary_Table_Row, LastCol + 4).Interior.ColorIndex = 4
                                    Else
                                    ws.Cells(Summary_Table_Row, LastCol + 4).Interior.ColorIndex = 3
                                    End If
                                   
                                ws.Cells(Summary_Table_Row, LastCol + 6).Value = Total_Volume
                                    ' Add one to the summary table row
                                    Summary_Table_Row = Summary_Table_Row + 1
     
                        ' Reset the Opening and Closing Price
                        Opening_Price = ws.Cells(i + 1, 3)
                        Closing_Price = 0
                        Price_Change = 0
                        Percentage_Change = 0
                        Total_Volume = 0

                        ' If the cell immediately following a row is the same ticker symbol...
                    Else

                    ' Add to the Total Volume Change
                   
                    Total_Volume = Total_Volume + ws.Cells(i, 7).Value
                   
                    End If
                                       
                Next i
               
    ' Find the greatest increase, greatest decrease and volume of the stock in (LastCol + 5) Column
               
       Dim G_INCREASE As Double
       G_INCREASE = 0
       Dim G_DECREASE As Double
       G_DECREASE = 0
       Dim G_Volume As Double
       G_Volume = 0
       Dim Ticker_max As String
       Dim Ticker_min As String
       Dim Ticker_Volmax As String
       
       
       
       ' Find the last row and last column in the summary table that contains data
          LastRow_Summary = ws.Cells(Rows.Count, LastCol + 3).End(xlUp).Row
         
         
       ' Create Summary Table 2
                Dim Summary_Table2_Row As Integer
                Summary_Table2_Row = 2
        'print Greatest % increase, greatest % decrease, volume in the summary table 2
                ws.Range("O2").Value = "Greatest % Increase"
                ws.Range("O3").Value = "Greatest % Decrease"
                ws.Range("O4").Value = "Greatest Volume"
                ws.Cells(1, "P").Value = "Ticker"
                ws.Cells(1, "Q").Value = "Value"
               
                 ' Loop to check the max and min
                 For i = 2 To LastRow_Summary
                 
                 
                    ' Check for the G_increase max
                        If ws.Cells(i, LastCol + 5).Value > G_INCREASE Then
                           G_INCREASE = ws.Cells(i, LastCol + 5).Value
                           Ticker_max = ws.Cells(i, 10).Value
                           
                        End If
                       
                     ' Check for min
                     
                        If ws.Cells(i, LastCol + 5).Value < G_DECREASE Then
                           G_DECREASE = ws.Cells(i, LastCol + 5).Value
                           Ticker_min = ws.Cells(i, 10).Value
                       
                        End If
                                               
                     ' find stock with max volume
                        If ws.Cells(i, 13).Value > G_Volume Then
                           G_Volume = ws.Cells(i, 13).Value
                           Ticker_Volmax = ws.Cells(i, 10).Value
                           
                        End If
                       
                  Next i
                'print the MIN AND MAX
                ws.Cells(2, 17).Value = FormatPercent(G_INCREASE, 3)
                ws.Cells(3, 17).Value = FormatPercent(G_DECREASE, 3)
                ws.Cells(4, 17).Value = G_Volume
                ws.Cells(2, 16).Value = Ticker_max
                ws.Cells(3, 16).Value = Ticker_min
                ws.Cells(4, 16).Value = Ticker_Volmax
 
Next ws

    MsgBox ("Work is Done")

End Sub
