Attribute VB_Name = "Module1"
Sub Stock_Loop_For_Each_Worksheet()

'Task
' Easy----loop through each year (tab) and sum the total volume each stock had and desplay it with the ticket sysbol
' Moderate---- yearly change in the stock. +/- from frist record to last for a stock.; Percentage of that change, color code, +=Green, -=Red
' Hard----Locate the stock with the "Greatest % increase", "Greatest % Decrease" and "Greatest total volume"

'Fields(results)
' Easy---- column I = "Ticker"; column  L = "Total Stock Volume"
' Moderate---- column J = "Yearly Change"; column K = "Percent Change"
' Hard---- column o = [winers]; P = "Ticker"; Q = "Value"

' reprt this onto the same sheet that that year's  data is on
Dim ws As Worksheet
Dim starting_ws As Worksheet
Set starting_ws = ActiveSheet 'remember which worksheet is active in the beginning

For Each ws In ThisWorkbook.Worksheets
    ws.Activate
    
    

'Set Up the tab (will need to be added to the loop(ar start of each tab))
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"

Cells(1, 16).Value = "Ticker"
Cells(1, 17).Value = "Value"
Cells(2, 15).Value = "Greatest % increase"
Cells(3, 15).Value = "Greatest % Decrease"
Cells(4, 15).Value = "Greatest total volume"


Dim sht As Worksheet
Dim LastRow As Long
Dim OutputRow As Long
Dim TotalStockValue As Long
Dim CurrentCellValue As Long

Dim StockOpen As Single
Dim StockClose As Single
Dim StockChange As Single
Dim StockPercentChange As Single

'set initial values for these variables
StockOpen = Cells(2, 3).Value
StockClose = 0
StockChange = 0
StockPercentChange = 0

OutputRow = 2
TotalStockVol = 0
CurrentCellValue = 0

'get the number for the last row in the data
Set sht = ActiveSheet
 LastRow = sht.Cells(sht.Rows.Count, "A").End(xlUp).Row
  'MsgBox (LastRow)
  
  
    
   For i = 2 To LastRow
  ' For i = 2 To 523
    
    ' Searches for next cell is different than current cell
    
    If Cells(i + 1, 1).Value = Cells(i, 1).Value Then
'CurrentCellValue = Cells(i, 7).Value
' MsgBox Cells(i, 7).Value
    TotalStockVol = Cells(i, 7).Value + TotalStockVol
            'fixes the $0 at start of year problem
            If StockOpen = 0 Then
            StockOpen = Cells(i, 3).Value
            End If
    
    ElseIf Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

      ' Message Box the value of the current cell and value of the next cell
      Cells(OutputRow, 9).Value = Cells(i, 1)
   TotalStockVol = Cells(i, 7).Value + TotalStockVol
      Cells(OutputRow, 12).Value = TotalStockVol
      StockClose = Cells(i, 6).Value
      StockChange = (StockClose - StockOpen)
      'fix the no stock that year issue
            If StockClose = 0 Then
            StockPercentChange = 0
            Else
            StockPercentChange = ((StockClose / StockOpen) - 1)
            End If
            
      

      
     'MsgBox (StockOpen)
      '    MsgBox (StockClose)
      '  MsgBox (StockChange)
      '    MsgBox (StockPercentChange)
      
       Cells(OutputRow, 10).Value = StockChange
       Cells(OutputRow, 11).Value = StockPercentChange
      
        ' nested if for color
            If StockChange >= 0 Then
            Cells(OutputRow, 10).Interior.ColorIndex = 4
            ElseIf StockChange < 0 Then
            Cells(OutputRow, 10).Interior.ColorIndex = 3
            End If
      
      OutputRow = OutputRow + 1
      TotalStockVol = 0
      StockOpen = Cells(i + 1, 3).Value
      StockClose = 0


    End If

  Next i

'format column K to be Percentage and go to two decimal places
    Columns("K:K").Select
    Selection.Style = "Percent"
    Selection.NumberFormat = "0.00%"
      
      
      
    'Start of the section that does the Greatest increase, decrease and total volume
      
      'finds the last row number for the summarized  data
      Dim LastRow2 As Long
     LastRow2 = sht.Cells(sht.Rows.Count, "I").End(xlUp).Row
     
'sets the first data to check against
    Dim GreatInc As Single
    Dim GreatDec As Single
    Dim GreatVol As Double
  GreatInc = Cells(2, 11).Value
  GreatDec = Cells(2, 11).Value
  GreatVol = Cells(2, 12).Value
  
'sets the ticker for the first data to check against
    Dim GreatIncT As String
    Dim GreatDecT As String
    Dim GreatVolT As String
  GreatIncT = Cells(2, 9).Value
  GreatDecT = Cells(2, 9).Value
  GreatVolT = Cells(2, 9).Value
     
     
     For i = 2 To LastRow2
    
    
    'ifs to find greatest...
    If Cells(i, 11).Value > GreatInc Then
        GreatInc = Cells(i, 11).Value
        GreatIncT = Cells(i, 9).Value
         End If
    If Cells(i, 11).Value < GreatDec Then
        GreatDec = Cells(i, 11).Value
        GreatDecT = Cells(i, 9).Value
         End If
    If Cells(i, 12).Value > GreatVol Then
        GreatVol = Cells(i, 12).Value
        GreatVolT = Cells(i, 9).Value
         End If

  Next i
  
'outputs the findings to the cells
Cells(2, 16).Value = GreatIncT
Cells(2, 17).Value = GreatInc
Cells(3, 16).Value = GreatDecT
Cells(3, 17).Value = GreatDec
Cells(4, 16).Value = GreatVolT
Cells(4, 17).Value = GreatVol

 'format column Q cells 2 and 3 to be Percentage and go to two decimal places
    Range("Q2:Q3").Select
    Selection.Style = "Percent"
    Selection.NumberFormat = "0.00%"
    
    
'resize all columns that the script added
    Columns("I:Q").EntireColumn.AutoFit


    ws.Cells(1, 1) = 1 'this sets cell A1 of each sheet to "1"
Next

starting_ws.Activate 'activate the worksheet that was originally active

     

End Sub

