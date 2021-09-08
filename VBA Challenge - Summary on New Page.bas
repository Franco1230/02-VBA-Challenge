Attribute VB_Name = "Module1"
Sub VBAchallenge():
Dim resulttable As Integer
Dim tickers As String
Dim totalvol As Double
Dim Summary_Table_Row As Integer
Dim openvaule As Double
Dim closevalue As Double
Dim yearlychange As Double
Dim changepercentage As Double
Dim greatdecrease As String
Dim greatincrease As String

Sheets.Add.Name = "Summary"
Sheets("Summary").Move Before:=Sheets(1)
    
Set combined_sheet = Worksheets("Summary")
Sheets("Summary").Activate
Range("A1").Value = "Ticker"
Range("B1").Value = "Yearly Change"
Range("C1").Value = "Percentage"
Range("D1").Value = "Total Volume"
Range("E1").Value = "From Work Sheet"
Summary_Table_Row = 2
resulttable = 1
Columns(resulttable).NumberFormat = "@"
Columns(resulttable + 2).NumberFormat = "0.00%"
Columns(resulttable + 3).NumberFormat = "#,##0 "
  
For Each WS In Worksheets

     If WS.Name <> "Summary" Then
  
     WS.Activate
     totalvol = 0
     lastrow = Cells(Rows.Count, 1).End(xlUp).Row
     closevalue = Cells(2, 6).Value
     openvalue = Cells(2, 3).Value

          For i = 2 To lastrow
    
               If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
               tickers = Cells(i, 1).Value
               totalvol = totalvol + Cells(i, 7).Value
               yearlychange = closevalue - openvalue
   
                    If yearlychange <> 0 Then

                    changepercentage = yearlychange / openvalue
    
                    Else
    
                    changepercentage = 0
    
                    End If

               Sheets("Summary").Cells(Summary_Table_Row, resulttable) = tickers
               Sheets("Summary").Cells(Summary_Table_Row, resulttable + 1) = yearlychange
               Sheets("Summary").Cells(Summary_Table_Row, resulttable + 2) = changepercentage
               Sheets("Summary").Cells(Summary_Table_Row, resulttable + 3) = totalvol
               Sheets("Summary").Cells(Summary_Table_Row, resulttable + 4) = ActiveSheet.Name
               Summary_Table_Row = Summary_Table_Row + 1
               totalvol = 0
               closevalue = Cells(i + 1, 6).Value
               openvalue = Cells(i + 1, 3).Value
        
               Else

                    If Cells(i + 1, 2).Value > Cells(i, 2).Value Then
   
                    closevalue = Cells(i + 1, 6).Value
   
                    ElseIf Cells(i + 1, 2).Value < Cells(i, 2).Value Then
   
                    openvalue = Cells(i + 1, 3).Value
   
                    End If

               If openvalue = 0 Then
                
                    openvalue = Cells(i + 1, 3).Value
                
               End If
   
               totalvol = totalvol + Cells(i, 7).Value

               End If
     
          Next i
  
     End If
  
Next WS
   
Sheets("Summary").Activate
Columns("C:C").Select

Selection.Style = "Percent"

Selection.NumberFormat = "0.00%"
Columns("D:D").Select

Selection.NumberFormat = "#,##0"
Columns("A:D").Select
Columns("A:E").EntireColumn.AutoFit
Range("B2:B" & Summary_Table_Row).Select

Selection.FormatConditions.Delete
Set condition1 = Range("B2:B" & Summary_Table_Row).FormatConditions.Add(xlCellValue, xlGreater, "=0")

Set condition2 = Range("B2:B" & Summary_Table_Row).FormatConditions.Add(xlCellValue, xlLess, "=0")

With condition1

     .Interior.ColorIndex = 4

End With

With condition2

     .Interior.ColorIndex = 3

End With

Range("G2").Value = "Greatest % Increase"
Range("G3").Value = "Greatest % Decrease"
Range("G4").Value = "Greatest Total Volume"
Range("H1").Value = Range("A1").Value
Range("I1").Value = "Value"
Range("J1").Value = "From Work Sheet"
Range("I2").Value = Application.WorksheetFunction.Max(Range("C:C"))

MaxRow = Application.WorksheetFunction.Match(Application.WorksheetFunction.Max(Range("C:C")), Range("C:C"), 0)
MinRow = Application.WorksheetFunction.Match(Application.WorksheetFunction.Min(Range("C:C")), Range("C:C"), 0)

Range("H2").Value = Cells(MaxRow, 1)
Range("I2").Style = "Percent"
Range("I2").NumberFormat = "0.00%"
Range("I3").Value = Application.WorksheetFunction.Min(Range("C:C"))
Range("H3").Value = Cells(MinRow, 1)
Range("J2").Value = Cells(MaxRow, 5)
Range("J3").Value = Cells(MinRow, 5)
Range("I3").Style = "Percent"
Range("I3").NumberFormat = "0.00%"
Range("I4").Value = Application.WorksheetFunction.Max(Range("D:D"))

MaxRow = Application.WorksheetFunction.Match(Application.WorksheetFunction.Max(Range("D:D")), Range("D:D"), 0)

Range("H4").Value = Cells(MaxRow, 1)
Range("J4").Value = Cells(MaxRow, 5)
Range("I4").NumberFormat = "#,##0"
Range("H1:J1,G2:G4").Activate
    
Selection.Font.Bold = True
Range("G1:J4").Select

Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    
Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    
With Selection.Borders(xlEdgeLeft)
        
     .Weight = xlThin
    
End With
    
With Selection.Borders(xlEdgeRight)
        
     .Weight = xlThin
    
End With

With Selection.Borders(xlEdgeTop)
        
     .Weight = xlThin
    
End With
    
With Selection.Borders(xlEdgeBottom)
        
     .Weight = xlThin
   
End With
    
With Selection.Borders(xlInsideVertical)
        
     .Weight = xlThin
    
End With
    
With Selection.Borders(xlInsideHorizontal)
        
     .Weight = xlThin
    
End With

Columns("G:J").EntireColumn.AutoFit

End Sub
