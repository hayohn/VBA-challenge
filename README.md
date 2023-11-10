# VBA-challenge
Module 2
the Folder has, screen shots. the VBA Excel file for both the actual data and testing.

below is the VBA code

Sub Test()


    Dim ws As Worksheet
    Dim lastRow As Long
    Dim ticker As String
    Dim open_Price As Double
    Dim close_Price As Double
    Dim yearly_Change As Double
    Dim percent_Change As Double
    Dim total_Vol As Double
    Dim greatest_Increase As Double
    Dim greatest_Decrease As Double
    Dim greatest_Vol As Double
    Dim greatest_IncreaseTicker As String
    Dim greatest_DecreaseTicker As String
    Dim greatest_VolTicker As String

    ' Initialize greatest_Vol to 0
    greatest_Vol = 0

    ' Loop through all worksheets in the active workbook
    For Each ws In ThisWorkbook.Worksheets
        ' Find the last row in the current worksheet
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        total_Vol = 0
        open_Price = ws.Cells(2, 3).Value

        ' Add the new column headers for the new inputs
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percentage Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"

        ' Loop through the rows (each row of data)
        For i = 2 To lastRow
            ticker = ws.Cells(i, 1).Value
            close_Price = ws.Cells(i, 6).Value
            total_Vol = total_Vol + ws.Cells(i, 7).Value

            ' To calculate the yearly and percent changes
            yearly_Change = close_Price - open_Price
            If open_Price <> 0 Then
                percent_Change = (yearly_Change / open_Price) * 100
            Else
                percent_Change = 0
            End If

            ' To add the information(inputs) into the needed cells
            ws.Cells(i, 9).Value = ticker
            ws.Cells(i, 10).Value = yearly_Change
            ws.Cells(i, 11).Value = percent_Change
            ws.Cells(i, 12).Value = total_Vol

            ' To update the open_Price for the next ticker
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
                open_Price = ws.Cells(i + 1, 3).Value
            End If

            If percent_Change > greatest_Increase Then
                greatest_Increase = percent_Change
                greatest_IncreaseTicker = ticker
            ElseIf percent_Change < greatest_Decrease Then
                greatest_Decrease = percent_Change
                greatest_DecreaseTicker = ticker
            End If

            If total_Vol > greatest_Vol Then
                greatest_Vol = total_Vol
                greatest_VolTicker = ticker
            End If
        Next i

        ' To apply the conditional formatting to view the results (green = 4 positive, red = 3 negative)
        For j = 2 To lastRow
            If ws.Cells(j, 10) > 0 Then
                ws.Cells(j, 10).Interior.ColorIndex = 4
            ElseIf ws.Cells(j, 10) < 0 Then
                ws.Cells(j, 10).Interior.ColorIndex = 3
            End If
        Next j

        ' Add the greatest % and total volume summary
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(2, 16).Value = greatest_IncreaseTicker
        ws.Cells(2, 17).Value = greatest_Increase
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(3, 16).Value = greatest_DecreaseTicker
        ws.Cells(3, 17).Value = greatest_Decrease
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(4, 16).Value = greatest_VolTicker
        ws.Cells(4, 17).Value = greatest_Vol

    Next ws

End Sub
