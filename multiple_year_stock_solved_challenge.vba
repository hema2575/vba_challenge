Attribute VB_Name = "Module11"
Sub totalStockvol()
Dim k As Double
Dim book As Workbook
Dim sheet As Worksheet
Dim EOYCloserowNum As Double
Dim totalVol As Double
Dim grtPrcntInc As Double
Dim grtPrcntDec As Double

MsgBox ("starting to find total volume of stock")


For Each book In Workbooks
For Each sheet In book.Worksheets
MsgBox (sheet.Name)
EOYCloserowNum = 1
 For k = 2 To sheet.Cells(Rows.Count, 1).End(xlUp).Row
    If sheet.Cells(k, 1).Value <> sheet.Cells(k - 1, 1).Value Then
       totalVolrow = Cells(Rows.Count, 12).End(xlUp).Row + 1
       Cells(totalVolrow, 9) = sheet.Cells(k, 1).Value ' ticker name -> I
       Cells(totalVolrow, 12) = sheet.Cells(k, 12).Value 'calculating & updating totalstock vol -> L
       Cells(totalVolrow, 13) = sheet.Cells(k, 3).Value 'annualOpen ->M
       BOYOpen = sheet.Cells(k, 3).Value 'collecting annualOpen value into a variable
       EOYCloserowNum = EOYCloserowNum + 1
   Else
       EOYCloserowNum = EOYCloserowNum + 1
       Cells(totalVolrow, 12) = Cells(totalVolrow, 12) + sheet.Cells(k, 7).Value 'calculating & updating totalstock vol -> L
       Cells(totalVolrow, 14).Value = sheet.Cells(EOYCloserowNum, 6).Value 'EOYClose ->N
       EOYClose = Cells(EOYCloserowNum, 6).Value 'collecting annualclose value into a variable
       yrlyChng = EOYClose - BOYOpen   'Calculating yearly change
       Cells(totalVolrow, 10).Value = yrlyChng
         If yrlyChng < 0 Then    'color coding for yearly change
            Cells(totalVolrow, 10).Interior.ColorIndex = 3
            Else
            Cells(totalVolrow, 10).Interior.ColorIndex = 4
         End If
            If BOYOpen <> 0 Then    'calculating percent change
               prcntChng = (yrlyChng / BOYOpen) * 100
               Cells(totalVolrow, 11).Value = prcntChng
            End If
   End If
 Next k
Next sheet
Next book
MsgBox ("Done. Now finding greatest stockVolume total, greatest % Increase and greatest % Decrease")


j = 2
    totalVol = Cells(j, 12).Value
    grtPrcntInc = Cells(j, 11).Value
    grtPrcntDec = Cells(j, 11).Value
    
    For j = 3 To Cells(Rows.Count, 12).End(xlUp).Row
         If Cells(j, 12).Value > totalVol Then
         totalVol = Cells(j, 12).Value
         ticker1 = Cells(j, 9).Value
         Else
         End If
         
         If Cells(j, 11).Value > grtPrcntInc Then
         grtPrcntInc = Cells(j, 11).Value
         ticker2 = Cells(j, 9).Value
         Else
         End If
         
         If Cells(j, 11).Value < grtPrcntDec Then
         grtPrcntDec = Cells(j, 11).Value
         ticker3 = Cells(j, 9).Value
         Else
         End If
    Next j
         Cells(4, 16).Value = ticker1
         Cells(4, 17).Value = totalVol

         Cells(2, 16).Value = ticker2
         Cells(2, 17).Value = grtPrcntInc
     
         Cells(3, 16).Value = ticker3
         Cells(3, 17).Value = grtPrcntDec
End Sub


       
           
