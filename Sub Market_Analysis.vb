Sub Market_Analysis()
    
    Dim Table_Header As Boolean
    Table_Header = False
    Dim Tickers As String
    Dim Open1 As Double
    Dim Close1 As Double
    Dim total As Double
    total = 0
    Dim wsheet As Worksheet
    
    Dim Close_Percent As Double
    
    Dim Ticker_Slot As Long
    
    
    Lastrow = Cells(Rows.Count, 1).End(xlUp).Row

    For Each wsheet In ThisWorkbook.Worksheets
    Tickers = wsheet.Cells(2, 1).Value
    Ticker_Slot = 2
    total = 0

        If Table_Header Then
            wsheet.Range("I1").Value = "Ticker"
            wsheet.Range("J1").Value = "Yearly Change"
            wsheet.Range("K1").Value = "Percent Change"
            wsheet.Range("L1").Value = "Total Stock Volume"
        Else
            Table_Header = True
        End If
    Open1 = wsheet.Cells(2, 3).Value
        For i = 2 To Lastrow

            If wsheet.Cells(i, 1).Value <> wsheet.Cells(i + 1, 1).Value Then
            Tickers = wsheet.Cells(i, 1).Value
            wsheet.Range("I" & Ticker_Slot).Value = Tickers

            
            Close1 = wsheet.Cells(i, 6).Value
            yearly_change = Close1 - Open1
            
            
            total = total + wsheet.Cells(i, 7).Value
                
                
                Close1 = wsheet.Cells(i, 6).Value
                wsheet.Range("j" & Ticker_Slot).Value = yearly_change
                
                
                If Open1 <> 0 Then
                Close_Percent = (yearly_change / Open1) * 100
                
                Else
                MsgBox ("For " & Tickers & ", Row " & CStr(i) & ": Open1 =" & Open1 & ". Fix <open> field manually and save the spreadsheet.")
                End If
                If (yearly_change > 0) Then
                    'Fill column with GREEN color - good
                    wsheet.Range("J" & Ticker_Slot).Interior.ColorIndex = 4
                ElseIf (yearly_change <= 0) Then
                    'Fill column with RED color - bad
                    wsheet.Range("J" & Ticker_Slot).Interior.ColorIndex = 3
                End If
            wsheet.Range("k" & Ticker_Slot).Value = (CStr(Close_Percent) & "%")
            wsheet.Range("L" & Ticker_Slot).Value = total
            yearly_change = 0
            Close1 = 0
            Open1 = wsheet.Cells(i + 1, 3).Value
            Ticker_Slot = Ticker_Slot + 1
            total = 0
            Else
            total = total + wsheet.Cells(i, 7).Value
            End If

            Next i
        Next wsheet

End Sub
