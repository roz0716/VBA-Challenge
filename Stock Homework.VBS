Range("J1").Value = "ticker"
Range("K1").Value = "year change"
Range("L1").Value = "percentage change"
Range("M1").Value = "total volume stock"
Range("02").Value = "greatest percentage increase"
Range("03").Value = "greatest percentage decrease"
Range("04").Value = "largest total volume"
Range("Q1").Value = Value
Range("P1").Value = "ticker"
Range("K:M").EntireColumn.AutoFit
Range("0:0").EntireColumn.AutoFit


Dim Ticker_Symbol As String
Dim Volume_Stock_Total As Double
    Volume_Stock_Total = 0

Dim Open_Date_Year As Boolean
Open_Date_Yar = True

Dim lastrow As Long

Dim Open_Price_Year, End_Price_Year, Change_yearly As Double

Dim Change_percent As Double

Dim Row_Table_Summary As Integer
Row_Table_Summary = 2

lastrow = Cells(Rows.Count, "A").End(x1Up).Row

For i = 2 To lastrow

If Cells(i + 1).Value <> Cells(i, 1).Value Then

Ticker_Symbol = Cells(i, 1).Value

total_stock_volume = total_stock_volume + Cells(i, 7).Value

Range("J" & summary_table_row).Value = Ticker_Symbol

Range("M" & summary_row_table).Value = total_stock_volume

End_year_price = Cells(i, 6).Value
Year_change = year_end_price - open_yearly_price
Range("K" & summary_table_row).Value = yearly_change

Percentage_change = (yearly_change / open_year_price) * 100
Range("L" & summary_table_row).Value = percent_change
Range("L" & summary_row_table).NumberFormatformat = "0.00%"

If Range("K" & summary_row_table).Value > 0 Then
Range("K" & summary_table_row).Interior.ColorIndex = 4
Else
Range("K" & summary_table_row).Interior.ColorIndex = 3
End If

summary_table_row = summary_table_row + 1

Tot_stock_volume = 0

Open_year_date = True

Else

total_stock_volume = toal_stock_volume + Cells(i, 7).Value

If Open_year_date And Cells(i, 3).Value <> 0 Then
open_year_price = Cells(i, 3).Value
Open_year_date = False

End If

Next i

Dim n As interger
Dim max_inc, min_dec As Long
Dim max_vol As Double
max_vol = 0
max_dec = 0
max_vol = 0

n = ws.Cells(Rows.Count, 11).End(x1Up).Row

For j = 2 To n

If Range("L" & j).Value > max_inc Then
max_inc = Range("L" & j).Value
Range("q2").Value = max_inc
Range("p2").Value = Range("3" & j).Value

Elself Range("L" & j).Value < min_dec then
 min_dec = Range("L" & j).Value
Range("Q3").Value = Range("L" & j).Value
Range("P3").Value = Range("J" & j).Value

End If

If Range("M" & j).Value >= max_vol Then
max_vol = Range("M" & j).Value
Range("Q4").Value = max_vol
Range("P4").Value = Range("j" & j).Value

End If

Range("Q2:Q3").NumberFormat = "0.00%"
Range("q:q").EntireColumn.AutoFit




Next j

Next ws

End Sub







 
























































