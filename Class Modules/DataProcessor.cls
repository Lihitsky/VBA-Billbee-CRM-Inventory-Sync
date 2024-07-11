Option Explicit

' Subroutine to insert data into the Excel worksheet
Public Sub InsertData(data As Collection)
    ' Get the worksheet to insert data into
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Storage Data")
    
    ' Clear the existing data in the worksheet
    ws.Cells.Clear

    ' Initialize row and column counters
    Dim i As Long
    Dim j As Long
    Dim lastRow As Long
    lastRow = 1

    ' Set the header row
    ws.Cells(lastRow, 1).Value = "SKU"
    ws.Cells(lastRow, 2).Value = "StockCurrent"
    
    ' Loop through the collection of data and insert it into the worksheet
    For i = 1 To data.Count
        For j = 1 To data(i)("Data").Count
            ' Get SKU and StockCurrent values from the data
            Dim SKU As Variant
            Dim StockCurrent As Variant
            SKU = data(i)("Data")(j)("SKU")
            StockCurrent = data(i)("Data")(j)("StockCurrent")
            lastRow = lastRow + 1
            
            ' Insert the values into the worksheet if they are not null
            If Not IsNull(SKU) And Not IsNull(StockCurrent) Then
                ws.Cells(lastRow, 1).Value = SKU
                ws.Cells(lastRow, 2).Value = StockCurrent
            End If
        Next j
    Next i
    
    ' Sort the data by SKU
    Dim dataRange As Range
    Set dataRange = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, 2))
    dataRange.Sort Key1:=ws.Cells(1, 1), Order1:=xlAscending, Header:=xlYes
End Sub
