Sub funcc():
    Dim totalsheets As Integer
    Dim ws As Worksheet
    Dim LastRow As Long
    Dim totalcount As Double
    Dim stockname As String
    Dim Summary_Table_Row As Integer
  
    Summary_Table_Row = 2
    totalcount = 0
    For Each ws In Worksheets
        
        LastRow = Cells(Rows.Count, 1).End(xlUp).Row
            For j = 2 To LastRow
            If ws.Cells(j + 1, "A").Value <> ws.Cells(j, "A").Value Then
                stockname = ws.Cells(j + 1, "A").Value
                totalcount = totalcount + ws.Cells(j, "G")
                ws.Range("K" & Summary_Table_Row).Value = stockname
                ws.Range("M" & Summary_Table_Row).Value = totalcount
                Summary_Table_Row = Summary_Table_Row + 1
                totalcount = 0
            Else
                  totalcount = totalcount + ws.Cells(j, "G").Value
              End If
             Next j
               totalcount = 0
               Summary_Table_Row = 2
                
    Next ws
End Sub
