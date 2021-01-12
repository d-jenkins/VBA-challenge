Sub stock_data()

  ' Set an initial variable for holding the Ticker Symbols
  Dim Ticker As String

  ' Set an initial variable for holding the yearly change
  Dim Yearly_Change As Double
  Yearly_Change = 0
  
  ' Set an initial variable for holding the volume
  Dim Total_Stock_Volume As Double
  Total_Stock_Volume = 0

  ' Keep track of the location for each Ticker Symbol in the summary table
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2

    ' Keep track of the location for the Year Change in the summary table
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2

  ' Loop through all Ticker data for a given year
  For i = 1 To 705713

    ' Check if we are still within the same ticker symbol, if it is not...
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

      ' Set the Ticker symbol
      Ticker = Cells(i, 1).Value
    
      ' Print the Ticker symbol in the Summary Table
      Range("I" & Summary_Table_Row).Value = Ticker

      ' Print the Year Change to the Summary Table
      Range("J" & Summary_Table_Row).Value = Yearly_Change
      
      ' Print the Total Stock Amount to the Summary Table
      Range("K" & Summary_Table_Row).Value = Total_Stock_Volume
      
      ' Reset Total Stock Volume
      Total_Stock_Volume = 0
      
      
      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1

    ' If the cell immediately following a row is the same brand...
    Else
    
      ' Calculate total volume
      Total_Stock_Volume = (Total_Stock_Volume + Cells(i, 7).Value) - 1

      ' Add to the Brand Total
      ' Brand_Total = Brand_Total + Cells(i, 3).Value
      '100 * (p2 - p1) / p
    

    End If

  Next i

End Sub







---------------------------

Sub stock_data()

  ' Set an initial variable for holding the Ticker Symbols
  Dim Ticker As String

  ' Set an initial variable for holding the yearly change
  Dim Yearly_Change As Double
  Yearly_Change = 0

  ' Keep track of the location for each Ticker Symbol in the summary table
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2

  ' Loop through all Ticker data for a given year
  For i = 2 To 43398

    ' Check if we are still within the same credit card brand, if it is not...
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

      ' Set the Ticker symbol
      Ticker = Cells(i, 1).Value

      ' Add to the Brand Total
      'Brand_Total = Brand_Total + Cells(i, 3).Value [Come back to this]
    

      ' Print the Credit Card Brand in the Summary Table
      Range("I" & Summary_Table_Row).Value = Ticker

      ' Print the Brand Amount to the Summary Table
      'Range("J" & Summary_Table_Row).Value = Yearly_Change [Come back to this]
    

      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset the Brand Total
      'Brand_Total = 0 [Come back to this]

    ' If the cell immediately following a row is the same brand...
    Else

      ' Add to the Brand Total
      Brand_Total = Brand_Total + Cells(i, 3).Value

    End If

  Next i

End Sub




-------------------------------------------

 ' Loop through all Ticker data for a given year
  For i = 1 To 705713

    ' Check if we are still within the same ticker symbol, if it is not...
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

      ' Set the Ticker symbol
      Ticker = Cells(i, 1).Value
      
      ' Set the Yearly Change
      'Yearly_Change = Cells(i, 2).Value
      
      'Set the Totatl Stock Volume
      'Total_Stock_Volume = Cells(i, 3).Value
    
      ' Print the Ticker symbol in the Summary Table
      Range("I" & Summary_Table_Row).Value = Ticker

      ' Print the Year Change to the Summary Table
      Range("J" & Summary_Table_Row).Value = Yearly_Change
      
      ' Print the Total Stock Amount to the Summary Table
      Range("K" & Summary_Table_Row).Value = Total_Stock_Volume
      
      ' Reset Total Stock Volume
      Total_Stock_Volume = 0
      
      
      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1

    ' If the cell immediately following a row is the same brand...
    Else
    
      ' Calculate total volume
      Total_Stock_Volume = (Total_Stock_Volume + Cells(i, 7).Value) - 1

      
      ' Add to the Brand Total
      ' Brand_Total = Brand_Total + Cells(i, 3).Value
      '100 * (p2 - p1) / p
    

    End If

  Next i

End Sub


-------------------------------------


Sub stock_data()
 ' Setting header values
 Range("I1:L1").Value = Array("ticker symbol", "total stock volume", "yearly change ($)", "percent change")
 
 ' Declare a variable for the worksheet
  'Dim ws As Worksheet

  ' Set an initial variable for headers the Ticker Symbols, Total Stock Volume
  Dim Ticker As String
  Dim Total_Stock_Volume As String
  Dim open_price As Double
  Dim Close_Price As Double
  Dim Yearly_Change As Double
  Dim Percent_Change As Double

  ' Keep track of the location for the summary table
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 1


  ' Set an initial variable for holding the volume
  'Dim Total_Stock_Volume_Amount As Double
  'Total_Stock_Volume_Amount = 0
  
  ' Set an initial variable for holding the open price
  'Dim open_price As Double
  
  ' Set an initial variable for holding the close price
  'Dim Close_Price As Double
  
  ' Set an initial variable for holding the yearly change
  'Dim Yearly_Change As Double

  ' Set an initial variable for holding the percent change
  'Dim Percent_Change As Double

  ' Loop through all Ticker data for a given year
  'For i = 1 To 705713

    ' Check if we are still within the same ticker symbol, if it is not...
    'If Cells(i + 1, 1).Value <> Cells(I, 1).Value Then

      ' Set the Ticker symbol
      'Ticker = Cells(I, 1).Value
      
      ' Print the Ticker symbol in the Summary Table
      'Range("I" & Summary_Table_Row).Value = Ticker
      
      ' Set the Total Stock Volume
      'Total_Stock_Volume = Cells(i, 2).Value
      
      ' Print the Total Stock Amount to the Summary Table
      'Range("K" & Summary_Table_Row).Value = Total_Stock_Volume_Amount
      
      ' Add one to the summary table row
      'Summary_Table_Row = Summary_Table_Row + 1
      
      ' Set the Yearly Change
      'Yearly_Change = Cells(i, 2).Value
      

    ' If the cell immediately following a row is the same brand...
    'Else
    
      ' Calculate total volume
      'Total_Stock_Volume_Amount = Total_Stock_Volume_Amount + Cells(I + 1, 7).Value
      
      'Reset the Total Stock Volume
      'Total_Stock_Volume = 0

      
      ' Add to the Brand Total
      ' Brand_Total = Brand_Total + Cells(i, 3).Value
      '100 * (p2 - p1) / p
    

    'End If

  'Next I

End Sub





----------------------------------------------

Sub stock_data()

 ' Setting header values
 Range("I1:L1").Value = Array("ticker symbol", "total stock volume", "yearly change ($)", "percent change")
 
 ' Declare a variable for the worksheet
  'Dim ws As Worksheet

  ' Name variable for headers and for plac
  Dim Ticker3 As String
  Dim Total_Stock_Volume As Long
  Dim open_price As Double
  Dim Close_Price As Double
  Dim Yearly_Change As Double
  Dim Percent_Change As Double

  ' Keep track of the location for the summary table
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 1


  ' Set an initial variable for holding the volume
  'Dim Total_Stock_Volume_Amount As Double
  'Total_Stock_Volume_Amount = 0
  
  ' Set an initial variable for holding the open price
  'Dim open_price As Double
  
  ' Set an initial variable for holding the close price
  'Dim Close_Price As Double
  
  ' Set an initial variable for holding the yearly change
  'Dim Yearly_Change As Double

  ' Set an initial variable for holding the percent change
  'Dim Percent_Change As Double

  ' Loop through all Ticker data for a given year
  For i = 1 To 705713

    ' Check if we are still within the same ticker symbol, if it is not...
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

      ' Set the Ticker symbol
      Ticker3 = Cells(i, 1).Value
      
      ' Print the Ticker symbol in the Summary Table
      Range("I" & Summary_Table_Row).Value = Ticker3
      
      ' Set the Total Stock Volume
      'Total_Stock_Volume = Cells(i, 2).Value
      
      ' Print the Total Stock Amount to the Summary Table
      'Range("K" & Summary_Table_Row).Value = Total_Stock_Volume_Amount
      
      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
      ' Set the Yearly Change
      'Yearly_Change = Cells(i, 2).Value
      

    ' If the cell immediately following a row is the same brand...
    Else
    
      ' Calculate total volume
      Total_Stock_Volume_Amount = Total_Stock_Volume_Amount + Cells(i + 1, 7).Value
      
      'Reset the Total Stock Volume
      'Total_Stock_Volume = 0

      
      ' Add to the Brand Total
      ' Brand_Total = Brand_Total + Cells(i, 3).Value
      '100 * (p2 - p1) / p
    

    End If

  Next i

End Sub

