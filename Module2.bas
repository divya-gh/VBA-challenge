Attribute VB_Name = "Module1"

Sub RunThroughOneYearStock()
'Delare the required variables'
  Dim ticker As String
  Dim lastrow As Long
  Dim Openprice As Variant
  Dim Closeprice As Variant
  Dim TotalVolume As Variant
  Dim Percentage As Variant
  Dim i As Long
  Dim j As Integer
  
  
  'Print the headers on cloumns I, J, K and L '
  ''-----------------------------------------''
  Range("I1").Value = "Ticker"
  Range("J1").Value = "Yearly Change"
  Range("K1").Value = "Percent Change"
  Range("L1").Value = "Total Stock Volume"
  
  'Find the last Row of the sheet'
  lastrow = ActiveSheet.UsedRange.SpecialCells(xlCellTypeLastCell).Row
  'MsgBox lastrow
  
  'Initialize the variables'
  ticker = ""
  Openprice = 0
  Closeprice = 0
  TotalVolume = 0
  Percentage = 0
  
  'Use variable j to lop through rows in column I, J , L'
  j = 2
  
  'Loop through rows from 1st to last row'
  
  For i = 2 To lastrow
  
    'Get unique tickers and place them in column I'
     If Cells(i, 1).Value <> ticker Then
        ticker = Cells(i, 1).Value
        Cells(j, 9).Value = ticker
        
        'Print the values for yearly change for previous ticker in the row (j-1) and column I'
        If i > 2 Then
           Cells(j - 1, 10).Value = (Closeprice - Openprice)
           
           'To avoid overflow error, make sure close price is not devided by a 0'
           'calculate the percentage'
           If Openprice <> 0 Then
              Percentage = ((Closeprice / Openprice) - 1)
              
           Else
              Percentage = 0
            
          End If
        
           'Print percentage and total Volume in the column K and row j-1'
           'Format the percentage to have 2 decimals with symbol'
           Percentage = Format(Percentage, "0.00%")
           Cells(j - 1, 11).Value = Percentage
           Cells(j - 1, 12).Value = TotalVolume
        
        End If
        
        'Find first OpenPrice and stock Volume for each unique tickers'
        Openprice = Cells(i, 3).Value
        TotalVolume = Cells(i, 7).Value
        
        j = j + 1
        
        
      Else
        
        'calculate the Total Volume by adding each correspoing row for ever ticker and get thier close price'
        Closeprice = Cells(i, 6).Value
        TotalVolume = TotalVolume + Cells(i, 7).Value
        
      End If
      
   Next i
      
  
  'Conditional formatting to highlight positive change in Green and Negetice change in Red'
  '----------------------------------------------------------------------------------------'
  Dim lrow As Long
  
  lrow = Cells(Rows.Count, 10).End(xlUp).Row  'Find the last row in column I'
  
  For i = 2 To lrow
    'find if value is < 0 '
    If (Cells(i, 10).Value < 0) Then
        Cells(i, 10).Interior.Color = RGB(255, 0, 0)
    Else
        Cells(i, 10).Interior.Color = RGB(0, 255, 0)
    End If
    
  Next i
  
  ''---------------------------------------------------''
  ''Bonus Question ''
  ''---------------------------------------------------''
  
  Dim GreatIncrease As Variant
  Dim GreatDecrease As Variant
  Dim GreatTotalVolume As Long
  
  'Assign first value to compare '
  GreatIncrease = Cells(2, 11).Value
  GreatDecrease = Cells(2, 11).Value
  
  'Inerate with for loop to compare'
  For Each cell In Range("K3:K" & lrow)
    If (cell.Value > GreatIncrease) Then
       
       GreatIncrease = cell.Value
       
    ElseIf (cell.Value < GreatDecrease) Then
       GreatDecrease = cell.Value
       
    End If
  
  Next cell
'print to check greatest and smallest value'
 GreatIncrease = Format(GreatIncrease, "0.00%")
 GreatDecrease = Format(GreatDecrease, "0.00%")
 
 MsgBox GreatIncrease
 MsgBox GreatDecrease
 
 
End Sub
