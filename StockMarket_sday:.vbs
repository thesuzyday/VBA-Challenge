Sub TickerSt():

'Set variables
Dim yearlychange As Double
Dim percentchange As Double
Dim stockvol As Long
Dim ticker As String
Dim openscore As Double
Dim closescore As Double

On Error Resume Next

LastRowJ = ws.Cells(Rows.Count, 10).End(x1Up).Row
LastRowK = ws.Cells(Rows.Count, 11).End(x1Up).Row

yearlychange = 0
percentagechange = 0
stockvol = 0
openscore = 0
closescore = 0

'location for tickers in summary table
Dim summaryrow As Integer
summaryrow = 2

'Update summary headers
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"

'Loop through tickers

For i = 2 To 759001

    'Check ticker before moving on and add to summary
    If Cells(i - 1, 1).Value <> Cells(i, 1).Value Then
    
        'Find first openscore
        openscore = Cells(i, 3).Value
        
    ElseIf Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                'define ticker
                ticker = Cells(i, 1).Value

              'Print ticker in Table
              Range("I" & summaryrow).Value = ticker
     
            'define closing score
             closescore = Cells(i, 6).Value
        
    
             'Calculate Yearly Change
              yearlychange = yearlychange + (closescore - openscore)
         
             'Print Yearly Change for Column J
                Range("J" & summaryrow).Value = yearlychange
        
              'Divide change by start for Percent changed for Column K
            percentagechange = (yearlychange / openscore)
            Range("K2:K3001").NumberFormat = "0.00%"

            'Adding to stock volume total
            stockvol = stockvol + Cells(i, 7).Value
                Range("L2:L3001").NumberFormat = "000000000"
            
            'Print Percent Changed in Column K
            Range("K" & summaryrow).Value = percentagechange

            'Print Total Stock Volume in Table on Column L
          Range("L" & summaryrow).Value = stockvol

   'reset counts for all stats/prepare for next row
    summaryrow = summaryrow + 1
    yearlychange = 0
    percentagechange = 0
    stockvol = 0

'if it's same ticker
Else

'Add to current ticker totals
stockvol = stockvol + Cells(i, 7).Value

    End If
    
Next i

'Color Code Yearly Change
Dim changecount As Range
Set changecount = Range("J2:J3001")
For Each cell In changecount
If cell.Value = "" Then
    cell.Interior.ColorIndex = 2
    End If
    
    If cell.Value < "0" Then
        cell.Interior.ColorIndex = 3
    End If
    
    If cell.Value > "0" Then
        cell.Interior.ColorIndex = 4
    End If
    
    If cell.Value = "0" Then
        cell.Interior.ColorIndex = 6
        End If
    Next
    
    'Conditional Formatting Reference: https://stackoverflow.com/questions/44588473/excel-vba-format-based-on-cell-value-greater-than-less-than-equal-to
    
End Sub
