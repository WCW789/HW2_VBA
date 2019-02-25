Sub stock_loop()

 For Each ws In Worksheets
 
  ws.Columns("J:L").ColumnWidth = 16
  ws.Columns("O").ColumnWidth = 20
  ws.Columns("Q").ColumnWidth = 12
 
  Dim Stock_Name As String

  Dim Stock_Total As Double
  Stock_Total = 0

  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2
  
  LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

  Dim opennumber As Double
  Dim closenumber As Double
  
  opennumber = ws.Cells(2, 3).Value
  
  ws.Range("I1").Value = "Ticker"
  ws.Range("J1").Value = "Yearly Change"
  ws.Range("K1").Value = "Percent Change"
  ws.Range("L1").Value = "Total Stock Volume"
  ws.Range("P1").Value = "Ticker"
  ws.Range("Q1").Value = "Value"
  
  ws.Range("O2").Value = "Greatest % Increase"
  ws.Range("O3").Value = "Greatest % Decrease"
  ws.Range("O4").Value = "Greatest Total Volume"
  
  Dim MaxValue
  Dim MinValue
  Dim MaxVolume
  
  Dim MaxName As String
  Dim lookupvalue1
  
  Dim MinName As String
  Dim lookupvalue2
    
  Dim VolName As String
  Dim lookupvalue3
    

  For i = 2 To LastRow

    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

      Stock_Name = ws.Cells(i, 1).Value
      closenumber = ws.Cells(i, 6).Value
      Stock_Total = Stock_Total + ws.Cells(i, 7).Value
      ws.Range("I" & Summary_Table_Row).Value = Stock_Name
      ws.Range("J" & Summary_Table_Row).Value = closenumber - opennumber

      If opennumber = 0 Then
        ws.Range("K" & Summary_Table_Row).Value = 0
      Else
        ws.Range("K" & Summary_Table_Row).Value = (closenumber / opennumber) - 1
      End If
      
      ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
      
        If ws.Range("K" & Summary_Table_Row).Value >= 0 Then
            ws.Range("K" & Summary_Table_Row).Interior.ColorIndex = 4
        ElseIf ws.Range("K" & Summary_Table_Row).Value < 0 Then
            ws.Range("K" & Summary_Table_Row).Interior.ColorIndex = 3
        End If
      
      ws.Range("L" & Summary_Table_Row).Value = Stock_Total
      Summary_Table_Row = Summary_Table_Row + 1
      opennumber = 0
      opennumber = ws.Cells(i + 1, 3).Value
      closenumber = 0
      Stock_Total = 0
       
      MaxValue = ws.Application.WorksheetFunction.Max(ws.Range("K2:K" & LastRow))
      ws.Range("Q2").Value = MaxValue
      ws.Range("Q2").NumberFormat = "0.00%"
      
      MinValue = ws.Application.WorksheetFunction.Min(ws.Range("K2:K" & LastRow))
      ws.Range("Q3").Value = MinValue
      ws.Range("Q3").NumberFormat = "0.00%"
    
      MaxVolume = ws.Application.WorksheetFunction.Max(ws.Range("L2:L" & LastRow))
      ws.Range("Q4").Value = MaxVolume
      
    Else
      Stock_Total = Stock_Total + ws.Cells(i, 7).Value
      
    End If
  
  Next i
  
   lookupvalue1 = ws.Range("Q2").Value
   MaxName = ws.Application.WorksheetFunction.Index(ws.Range("I2:I" & LastRow), ws.Application.WorksheetFunction.Match(lookupvalue1, ws.Range("K2:K" & LastRow), 0))
   ws.Range("P2").Value = MaxName
   
   lookupvalue2 = ws.Range("Q3").Value
   MinName = ws.Application.WorksheetFunction.Index(ws.Range("I2:I" & LastRow), ws.Application.WorksheetFunction.Match(lookupvalue2, ws.Range("K2:K" & LastRow), 0))
   ws.Range("P3").Value = MinName
   
   lookupvalue3 = ws.Range("Q4").Value
   VolName = ws.Application.WorksheetFunction.Index(ws.Range("I2:I" & LastRow), ws.Application.WorksheetFunction.Match(lookupvalue3, ws.Range("L2:L" & LastRow), 0))
   ws.Range("P4").Value = VolName
   
Next ws

End Sub