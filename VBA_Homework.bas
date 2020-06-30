Attribute VB_Name = "Module1"
Sub Eval_Sheet()
'How many worksheets in workbook?
Dim ws As Worksheet
'SheetTotal = Application.Sheets.Count
Dim SheetCurrent As Integer

'   Loop through each sheet
    For Each ws In Worksheets
'       Loop through all tickers to get full list
        Dim Range_Start As Long
        Range_Start = 2
        
        Dim Range_End As Long
        Range_End = ws.Cells(Rows.Count, 1).End(xlUp).Row
'       Row Index Variable
        Dim Start_Row As Long
        Start_Row = 2
        Dim End_Row
        End_Row = 3
        
        Dim Ticker_Code As String
        Dim Ticker_Code_End As String
        
        Dim total_volume As Double
        total_volume = 0
        
        Dim Close_Price As Double
        Dim Open_Price As Double
        Dim Yearly_Change As Double
'       Start Summary column
        Dim Summary_Row As Long
        Summary_Row = 2
        Dim Summary_Column
        Summary_Column = 9
'        Summary Headers
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Volume"
    '       Loop through list on sheet
        For Ticker_Range = 2 To Range_End

                If ws.Cells(Ticker_Range, 1).Value <> ws.Cells(Ticker_Range + 1, 1).Value Then
                    ' Set current ticker to variable.
                    Ticker_Code = ws.Cells(Ticker_Range, 1).Value
                    ' Set current summary table row to ticker name.
                    ws.Cells(Summary_Row, 9).Value = Ticker_Code
                    
                    total_volume = total_volume + ws.Cells(Ticker_Range, 7).Value
                    ws.Cells(Summary_Row, 12).Value = total_volume
                    
                    total_volume = 0
                    
                    Close_Price = ws.Cells(Ticker_Range, 6).Value
                    Open_Price = ws.Cells(Start_Row, 3).Value
                    Yearly_Change = Close_Price - Open_Price
                    ws.Cells(Summary_Row, 10).Value = Yearly_Change
                    
                    If Open_Price = 0 Then
                        ws.Cells(Summary_Row, 11).Value = "%" & 0
                    Else
                        ws.Cells(Summary_Row, 11).Value = "%" & Round((Yearly_Change / Open_Price) * 100, 2)
                    End If

'                    format cells
                    If ws.Cells(Summary_Row, 10).Value < 0 Then
                    ws.Cells(Summary_Row, 10).Interior.ColorIndex = 3
                    Else
                    ws.Cells(Summary_Row, 10).Interior.ColorIndex = 4
                    End If
'                   go to next ticker code
                    Start_Row = Ticker_Range + 1
                    ' Increment Summary Table Row.
                    Summary_Row = Summary_Row + 1
                    
                Else
                    total_volume = total_volume + ws.Cells(Ticker_Range, 7).Value
                    
                End If
        Next Ticker_Range
        
'       Find Greatest values
        Dim Summary_Range As Integer
        Summary_Range = 2
        Dim Summary_End As Long
        Summary_End = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        Dim Percent_Up As Double
        Percent_Up = ws.Cells(2, 11).Value
        
        Dim Percent_Down As Double
        Percent_Down = ws.Cells(2, 11).Value
        
        Dim Total As LongLong
        Total = ws.Cells(2, 12).Value
        
        Dim Ticker_Up As String
        Ticker_Up = ws.Cells(2, 9).Value
        
        Dim Ticker_Down As String
        Ticker_Down = ws.Cells(2, 9).Value
        
        Dim Ticker_Total As String
        Ticker_Total = ws.Cells(2, 9).Value
        
'       Header and row titles
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        
        For Greatest = 2 To Summary_End
            If Percent_Up >= ws.Cells(Greatest, 11).Value Then
            Percent_Up = Percent_Up
            Ticker_Up = Ticker_Up
            Else
            Percent_Up = ws.Cells(Greatest, 11).Value
            Ticker_Up = ws.Cells(Greatest, 9).Value
            End If
            
            If Percent_Down <= ws.Cells(Greatest, 11).Value Then
            Percent_Down = Percent_Down
            Ticker_Down = Ticker_Down
            Else
            Percent_Down = ws.Cells(Greatest, 11).Value
            Ticker_Down = ws.Cells(Greatest, 9).Value
            End If
            
            If Total >= ws.Cells(Greatest, 12).Value Then
            Total = Total
            Ticker_Total = Ticker_Total
            Else
            Total = ws.Cells(Greatest, 12).Value
            Ticker_Total = ws.Cells(Greatest, 9).Value
            End If
            
        Next Greatest
        
        ws.Cells(2, 17).Value = "%" & Round(Percent_Up * 100, 2)
        ws.Cells(3, 17).Value = "%" & Round(Percent_Down * 100, 2)
        ws.Cells(4, 17).Value = Total
        ws.Cells(2, 16).Value = Ticker_Up
        ws.Cells(3, 16).Value = Ticker_Down
        ws.Cells(4, 16).Value = Ticker_Total

        
    Next ws
End Sub
