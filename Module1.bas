Attribute VB_Name = "Module112"
'Create a sub function to loop through the ticker and calculate its corresponding yearly change , percent change and Total stock volume'
'Calculations used :- 1. yearly change = closeprice -Openprice'
                    '2. Percent change = (closeprice-openprice)/openprice )* 100'
                    '3. Total volume = sum of volumes for individual tickers'
                    
Sub RunThroughOneYearStock()
'Delare the required variables'
  Dim ticker As String
  Dim lastrow As Long
  Dim Openprice As Variant
  Dim Closeprice As Variant
  Dim Totalvolume As Variant
  Dim Percentage As Variant
  Dim i As Long
  Dim j As Integer
  
  ''-----------------------------------------''
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
  Totalvolume = 0
  Percentage = 0
  
  'Use variable j to lop through rows in column I, J , L'
  j = 2
  
  'Loop through rows from 1st to last row'
  
  For i = 2 To lastrow
  
    'Get unique tickers and place them on column I'
     If Cells(i, 1).Value <> ticker Then
        ticker = Cells(i, 1).Value
        Cells(j, 9).Value = ticker
        
    'Call the function to Print on worksheet only for row3 and above'
        
        If i > 2 Then
        
        Call PrintOnWorksheet(j, Closeprice, Openprice, Percentage, Totalvolume)
                
        End If
        
    'Find the first OpenPrice and stock Volume for each unique tickers'
        Openprice = Cells(i, 3).Value
        Totalvolume = Cells(i, 7).Value
        
        j = j + 1
        
        
      Else
        
       'Calculate Total Volume and get the last closeprice in the variable'
        Closeprice = Cells(i, 6).Value
        Totalvolume = Totalvolume + Cells(i, 7).Value
        
        
        'Find if the closeprice value is from lastrow'
        If (Cells(lastrow, 6).Value = Closeprice) Then
           
     'Call the function to print the last row value'
           Call PrintOnWorksheet(j, Closeprice, Openprice, Percentage, Totalvolume)
           
           
        
        
        End If
     End If
      
   Next i
   
      
  ''--------------------------------------------------------------------------------------''
  'Conditional formatting to highlight positive change in Green and Negetice change in Red'
  ''---------------------------------------------------------------------------------------''
  
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
  Dim GreatTotalVolume As Variant
  Dim TickerIncrease As String
  Dim TickerDecrease As String
  
  'Assign first value to compare '
  GreatIncrease = Cells(2, 11).Value
  GreatDecrease = Cells(2, 11).Value
  
  'Inerate with for loop to compare'
  For Each cell In Range("K3:K" & lrow)
    If (cell.Value > GreatIncrease) Then
       
       GreatIncrease = cell.Value
       TickerIncrease = cell.Offset(, -2).Value
       
    ElseIf (cell.Value < GreatDecrease) Then
       GreatDecrease = cell.Value
       TickerDecrease = cell.Offset(, -2).Value
       
    End If
  
  Next cell
  
'Format to 2 decimals and print to check greatest and smallest value'
'--------------------------------'
 GreatIncrease = Format(GreatIncrease, "0.00%")
 GreatDecrease = Format(GreatDecrease, "0.00%")
 
 'MsgBox GreatIncrease
 'MsgBox GreatDecrease
 
 
 ''--------------------------------''
 'Print lables on worksheet'
 ''------------------------------''
 Cells(2, 15).Value = "Greatest % increase"
 Cells(3, 15).Value = "Greatest % decrease"
 Cells(4, 15).Value = "Greatest total volume"
 Cells(1, 16).Value = "Ticker"
 Cells(1, 17).Value = "Value"
 

 'Call function to calculate greatest total stock volume'
 '----------------------------------'
 GreatTotalVolume = GetTotalVolume(lrow)
 'MsgBox GreatTotalVolume'


 ''------------------------------------------------------''
 'Print Great%Increase , Great%Decrease and GreatTotalVolume on worksheet'
 ''------------------------------------------------------''
  Range("P2").Value = TickerIncrease
  Range("P3").Value = TickerDecrease
  Range("Q2").Value = GreatIncrease
  Range("Q3").Value = GreatDecrease
  Range("Q4").Value = GreatTotalVolume


 'Autofit text on column O and I through L '
 'reference :- from https://www.automateexcel.com/vba/autofit-columns-rows/'
 '---------------------------'
    
  Columns("O").EntireColumn.AutoFit
  Columns("I:L").EntireColumn.AutoFit
  
  
  ''---------------------------------------------------------------------------------------''
  'Additional Feature Just for fun!'
  ''---------------------------------------------------------------------------------------''
  'Create a table "Growth_Table" for range("O1:Q4")'
  'Code reference https://www.automateexcel.com/vba/tables-and-listobjects/'
  ''------------------------------------------------------------------------''
  Dim tablename As String
  Dim TableExists As Boolean
  
  'tablename = "Growth_Table"
  
  TableExists = False
On Error GoTo Skip
If ActiveSheet.ListObjects("Growth_Table").Name = "Growth_Table" Then
TableExists = True
End If
Skip:
    On Error GoTo 0
     
     If Not TableExists And (Range("O2").Value = "Greatest % increase") Then
     
        ActiveSheet.ListObjects.Add(xlSrcRange, Range("O1:Q4"), , xlYes).Name = "Growth_Table"
        ActiveSheet.ListObjects("Growth_Table").TableStyle = "TableStyleLight9"
     
     Else
       Exit Sub
     
     End If
    
 
End Sub


''--------------------------------------------------''
'Function to calculate GreatTotalvolume'
''--------------------------------------------------''

Private Function GetTotalVolume(lrow) As Variant
  Dim Totalvolume As Variant
  Dim TotalVolumeTIcker As String
  Totalvolume = Cells(2, 12).Value
    
  For i = 3 To lrow
    If Cells(i, 12).Value > Totalvolume Then     'If condition is met, then pass the greater value to Totalvolume'
       Totalvolume = Cells(i, 12).Value
       TotalVolumeTIcker = Cells(i, 12).Offset(, -3).Value  'Pass its ticker name to TickerTotalVolume variable'
    End If
       
  Next i
  
  'Return GetTotalVolume to main sub and print its ticker name on worksheet on rangeP4'
  ''------------------------------''
  GetTotalVolume = Totalvolume
  Range("P4").Value = TotalVolumeTIcker
  'MsgBox TotalVolumeTIcker
      
End Function


''----------------------------------------------------------------------------------''
'Function to Print Yearly Change, Percent Change and Total Stock Volume on Worksheet'
''----------------------------------------------------------------------------------''


Function PrintOnWorksheet(j As Integer, Closeprice As Variant, Openprice As Variant, Percentage As Variant, Totalvolume As Variant)
         
        'Print Yearly Change on Column J for previous ticker'
           Cells(j - 1, 10).Value = (Closeprice - Openprice)
           
        'To avoid overflow error, make sure close price is not devided by a 0'
        'Calculate the percentage'
           If Openprice <> 0 Then
              Percentage = ((Closeprice / Openprice) - 1)
              
           Else
              'Percentage = "N/A" ; keep percentage = 0 to avoid error while calculating great%increase '
               Percentage = 0
           End If
        
        'Format the percentage to have 2 decimals with symbol'
           Percentage = Format(Percentage, "0.00%")
           
        'Print percentage and Total Volume in the column K & L for previous ticker'
           
           Cells(j - 1, 11).Value = Percentage
           Cells(j - 1, 12).Value = Totalvolume



End Function




