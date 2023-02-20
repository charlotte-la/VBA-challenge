Attribute VB_Name = "Module1"
Sub WorksheetAnalysis()

Dim ws As Worksheet

For Each ws In Worksheets

Dim Ticker As String
Dim TableRow As Integer
Dim Row_Length As Double
Dim Year_Close As Double
Dim Year_Open As Double
Dim Yearly_Change As Double
Dim Vol As Double


    ' Adding column headers
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percentage Change"
    ws.Range("L1").Value = "Total Stock Volume"
    

TableRow = 2

' Finding the last row with data
Row_Length = ws.Cells(Rows.Count, 1).End(xlUp).Row


For i = 2 To Row_Length
    
    If ws.Cells(i, 3).Value = 0 Then
    
        If ws.Cells(i + 1).Value <> ws.Cells(i, 1).Value Then
            Ticker = ws.Cells(i, 1).Value
            
        End If

    ElseIf ws.Cells(i + 1, 1).Value = ws.Cells(i, 1).Value Then

        Vol = Vol + ws.Cells(i, 7).Value
        
        If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
            Year_Open = ws.Cells(i, 3).Value
        End If
        
    Else
    
     ' Store the current ticker and add the current volume
        Ticker = ws.Cells(i, 1).Value
        Vol = Vol + ws.Cells(i, 7).Value
        Year_Close = ws.Cells(i, 6).Value
        
        ws.Range("I" & TableRow).Value = Ticker
        ws.Range("L" & TableRow).Value = Vol
        
        If Vol > 0 Then
        
            ws.Range("J" & TableRow).Value = Year_Close - Year_Open
         
  
  ' Coloring the change cells green for positive change and red for negative change
            If ws.Range("J" & TableRow).Value > 0 Then
                ws.Range("J" & TableRow).Interior.ColorIndex = 4
                
            Else
                ws.Range("J" & TableRow).Interior.ColorIndex = 3
            
            End If
             
   ' Calculating the percent change and adding it to the table
        ws.Range("K" & TableRow).Value = ws.Range("J" & TableRow) / Year_Open
        
        Else
        
        ws.Range("J" & TableRow) = 0
        ws.Range("K" & TableRow) = 0
        
    
    End If
    
        ws.Range("K" & TableRow).Style = "percent"
        
        Vol = 0
        TableRow = TableRow + 1
        
End If
    
Next i
        
        

Next ws

End Sub
