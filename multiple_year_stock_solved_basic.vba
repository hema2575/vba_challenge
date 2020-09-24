Attribute VB_Name = "Module11"
Sub totalStockvol()
Dim k As Double
Dim book As Workbook
Dim sheet As Worksheet
Dim EOYCloserowNum As Double

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
MsgBox ("Done")
End Sub




        
           
