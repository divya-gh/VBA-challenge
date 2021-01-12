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
  
  Dim st As Worksheet
  
  Set st = ActiveSheet

  
  For Each ws In ThisWorkbook.Worksheets
      ws.Activate
      'On Error Resume Next
     
      ''-----------------------------------------''
      'Print the headers on cloumns I, J, K and L '
      ''-----------------------------------------''
      ws.Range("I1").value = "Ticker"
      ws.Range("J1").value = "Yearly Change"
      ws.Range("K1").value = "Percent Change"
      ws.Range("L1").value = "Total Stock Volume"
      
      
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
         If ws.Cells(i, 1).value <> ticker Then
            ticker = ws.Cells(i, 1).value
            ws.Cells(j, 9).value = ticker
            
        'Call the function to Print on worksheet only for row3 and above'
            
            If i > 2 Then
            
            Call PrintOnWorksheet(j, Closeprice, Openprice, Percentage, Totalvolume)
                    
            End If
            
        'Find the first OpenPrice and stock Volume for each unique tickers'
            Openprice = ws.Cells(i, 3).value
            Totalvolume = ws.Cells(i, 7).value
            
            j = j + 1
            
            
          Else
            
           'Calculate Total Volume and get the last closeprice in the variable'
            Closeprice = ws.Cells(i, 6).value
            Totalvolume = Totalvolume + ws.Cells(i, 7).value
            
            
            'Find if the closeprice value is from lastrow'
            If (ws.Cells(lastrow, 6).value = Closeprice) Then
               
         'Call the function to print the last row value'
               Call PrintOnWorksheet(j, Closeprice, Openprice, Percentage, Totalvolume)
               
               
            
            
            End If
         End If
          
       Next i
       
          
      ''--------------------------------------------------------------------------------------''
      'Conditional formatting to highlight positive change in Green and Negetice change in Red'
      ''---------------------------------------------------------------------------------------''
      
      Dim lrow As Long
      
      lrow = ws.Cells(Rows.Count, 10).End(xlUp).Row  'Find the last row in column I'
      
      For i = 2 To lrow
        'find if value is < 0 '
        If (ws.Cells(i, 10).value < 0) Then
            ws.Cells(i, 10).Interior.Color = RGB(255, 0, 0)
        Else
            ws.Cells(i, 10).Interior.Color = RGB(0, 255, 0)
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
      GreatIncrease = ws.Cells(2, 11).value
      GreatDecrease = ws.Cells(2, 11).value
      
      'Inerate with for loop to compare'
      For Each cell In Range("K3:K" & lrow)
        If IsNumeric(cell.value) And (cell.value > GreatIncrease) Then
           
           GreatIncrease = cell.value
           TickerIncrease = cell.Offset(, -2).value
           
        ElseIf IsNumeric(cell.value) And (cell.value < GreatDecrease) Then
           GreatDecrease = cell.value
           TickerDecrease = cell.Offset(, -2).value
           
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
     ws.Cells(2, 15).value = "Greatest % increase"
     ws.Cells(3, 15).value = "Greatest % decrease"
     ws.Cells(4, 15).value = "Greatest total volume"
     ws.Cells(1, 16).value = "Ticker"
     ws.Cells(1, 17).value = "Value"
     
    
     'Call function to calculate greatest total stock volume'
     '----------------------------------'
     GreatTotalVolume = GetTotalVolume(lrow)
     'MsgBox GreatTotalVolume'
    
    
     ''------------------------------------------------------''
     'Print Great%Increase , Great%Decrease and GreatTotalVolume on worksheet'
     ''------------------------------------------------------''
      ws.Range("P2").value = TickerIncrease
      ws.Range("P3").value = TickerDecrease
      ws.Range("Q2").value = GreatIncrease
      ws.Range("Q3").value = GreatDecrease
      ws.Range("Q4").value = GreatTotalVolume
    
    
     'Autofit text on column 'O' '
     'reference :- from https://www.automateexcel.com/vba/autofit-columns-rows/'
     '---------------------------'
        
      Columns("I:O").EntireColumn.AutoFit
      
    
         
         
   Next ws
   st.Activate
 
End Sub


''--------------------------------------------------''
'Function to calculate GreatTotalvolume'
''--------------------------------------------------''

Private Function GetTotalVolume(lrow) As Variant
  Dim Totalvolume As Variant
  Dim TotalVolumeTIcker As String
  Totalvolume = Cells(2, 12).value
    
  For i = 3 To lrow
  
  'If condition is met, then pass the greater value to Totalvolume'
    If (Cells(i, 12).value > Totalvolume) Then
       Totalvolume = Cells(i, 12).value
       TotalVolumeTIcker = Cells(i, 12).Offset(, -3).value  'Pass its ticker name to TickerTotalVolume variable'
    
    End If
       
  Next i
  
  'Return GetTotalVolume to main sub and print its ticker name on worksheet on rangeP4'
  ''------------------------------''
  GetTotalVolume = Totalvolume
  Range("P4").value = TotalVolumeTIcker
  'MsgBox TotalVolumeTIcker
      
End Function


''----------------------------------------------------------------------------------''
'Function to Print Yearly Change, Percent Change and Total Stock Volume on Worksheet'
''----------------------------------------------------------------------------------''


Function PrintOnWorksheet(j As Integer, Closeprice As Variant, Openprice As Variant, Percentage As Variant, Totalvolume As Variant)
         
        'Print Yearly Change on Column J for previous ticker'
           Cells(j - 1, 10).value = (Closeprice - Openprice)
           
        'To avoid overflow error, make sure close price is not devided by a 0'
        'Calculate the percentage'
           If Openprice <> 0 Then
              Percentage = ((Closeprice / Openprice) - 1)
              
           Else
              'If Open price is 0 then Percentage is NA '
               Percentage = "N/A"
           End If
        
        'Format the percentage to have 2 decimals with symbol'
           Percentage = Format(Percentage, "0.00%")
           
        'Print percentage and Total Volume in the column K & L for previous ticker'
           
           Cells(j - 1, 11).value = Percentage
           Cells(j - 1, 12).value = Totalvolume



End Function




