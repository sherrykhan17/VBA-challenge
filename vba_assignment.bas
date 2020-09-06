Sub run_all_sheets()
    Dim xSh As Worksheet
    Application.ScreenUpdating = False
    For Each xSh In Worksheets
        xSh.Select
        Call stock_market
    Next
    Application.ScreenUpdating = True
End Sub

Sub stock_market()
    Dim yearly_open As Double
    Dim yearly_close As Doublea
    Dim ticker As String
    Dim next_ticker As String
    Dim yearly_change_percent As Double
    Dim total_stock_vol As Double
    Dim current_row As Long
    Dim yearly_change As Double

    Dim max_increase(1)
    Dim min_increase(1)
    Dim max_total_vol(1)
a
    yearly_open = 0
    total_stock_vol = 0
    current_row = 1
    yearly_change = 0

    max_increase(0) = 0
    min_increase(0) = 0
    max_total_vol(0) = 0

    ' Define cell count as constants
    Const ticker_cell As Integer = 9
    Const yearly_change_cell As Integer = 10
    Const percent_change As Integer = 11
    Const yearly_vol_cell As Integer = 12
    Const challenge_cell_start As Integer = 13
    
    ' Add cell headers
    Cells(current_row, ticker_cell).Value = "Ticker"
    Cells(current_row, yearly_change_cell).Value = "Yearly Change"
    Cells(current_row, percent_change).Value = "Percent Change"
    Cells(current_row, yearly_vol_cell).Value = "Total Stock Volume"
    Cells(current_row, yearly_vol_cell).Value = "Total Stock Volume"
    Cells(current_row, challenge_cell_start + 2).Value = "Ticker"
    Cells(current_row, challenge_cell_start + 3).Value = "Value"

    ' Loop through each row
    For i = 2 To Rows.Count
        ticker = Cells(i, 1)
        next_ticker = ""
        On Error Resume Next
            next_ticker = Cells(i + 1, 1)
        total_stock_vol = total_stock_vol + Cells(i, 7)

        If yearly_open = 0 Then
            yearly_open = Cells(i, 3)
        End If

        ' End of ticker for a year
        If next_ticker <> ticker Then
            yearly_close = Cells(i, 6)
            yearly_change = yearly_close - yearly_open
            
            If yearly_open <> 0 Then
                yearly_change_percent = (yearly_change / yearly_open) * 100
            Else
                yearly_change_percent = 0
            End If

            ' Max increase
            If yearly_change_percent > max_increase(0) Then
                max_increase(0) = yearly_change_percent
                max_increase(1) = ticker
            End If

            ' Min increase
            If yearly_change_percent < min_increase(0) Then
                min_increase(0) = yearly_change_percent
                min_increase(1) = ticker
            End If

            ' Max volume
            If total_stock_vol > max_total_vol(0) Then
                max_total_vol(0) = total_stock_vol
                max_total_vol(1) = ticker
            End If

            ' Display results
            current_row = current_row + 1
            Cells(current_row, ticker_cell).Value = ticker
            Cells(current_row, yearly_change_cell).Value = Round(yearly_change, 2)
            Cells(current_row, percent_change).Value = Round(yearly_change_percent, 2)
            Cells(current_row, yearly_vol_cell).Value = total_stock_vol
            
            ' Apply background color
            If yearly_change > 0 Then
                Cells(current_row, yearly_change_cell).Interior.ColorIndex = 4
            ElseIf yearly_change < 0 Then
                Cells(current_row, yearly_change_cell).Interior.ColorIndex = 3
            End If

            ' Reset
            total_stock_vol = 0
            yearly_open = 0
        End If
    Next

    ' Print max / min infos
    Cells(2, challenge_cell_start + 1).Value = "Greatest % increase"
    Cells(2, challenge_cell_start + 2).Value = max_increase(1)
    Cells(2, challenge_cell_start + 3).Value = Round(max_increase(0), 2)
    Cells(3, challenge_cell_start + 1).Value = "Greatest % decrease"
    Cells(3, challenge_cell_start + 2).Value = min_increase(1)
    Cells(3, challenge_cell_start + 3).Value = Round(min_increase(0), 2)
    Cells(4, challenge_cell_start + 1).Value = "Greatest total volume"
    Cells(4, challenge_cell_start + 2).Value = max_total_vol(1)
    Cells(4, challenge_cell_start + 3).Value = max_total_vol(0)
End Sub
