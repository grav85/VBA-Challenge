Attribute VB_Name = "Module1"
Sub Ticker()
 
 ' Set Ws as a worksheet.
    Dim Ws As Worksheet
  
    ' Loop Worksheets.
    For Each Ws In Worksheets
    
        ' Variables and Values.
        Dim Ticker_Name As String
        Dim Total_Ticker_Volume As Double
        Dim Open_Price As Double
        Dim Close_Price As Double
        Dim Yearly_Change As Double
        Dim Percent_Change As Double
        Dim Summary_Table_Row As Long
        Dim Lastrow As Long
        Dim i As Long

        Ticker_Name = " "
        Total_Ticker_Volume = 0
        Open_Price = Ws.Cells(2, 3).Value
        Close_Price = 0
        Yearly_Change = 0
        Percent_Change = 0
        Summary_Table_Row = 2
        Lastrow = Ws.Cells(Rows.Count, 1).End(xlUp).Row

        ' Headers
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Yearly Change"
        Cells(1, 11).Value = "Percent Change"
        Cells(1, 12).Value = "Total Stock Volume"

               
        ' Loop
        For i = 2 To Lastrow
        
            'Loop Ticker Name
            If Ws.Cells(i + 1, 1).Value <> Ws.Cells(i, 1).Value Then
            
                ' Set Ticker Name
                Ticker_Name = Ws.Cells(i, 1).Value
                
                'Yearly_Change and Percent_Change
                Close_Price = Ws.Cells(i, 6).Value
                Yearly_Change = Close_Price - Open_Price

                If Open_Price <> 0 Then
                    Percent_Change = (Yearly_Change / Open_Price) * 100
                End If
                
                ' Sum of Total Volume
                Total_Ticker_Volume = Total_Ticker_Volume + Ws.Cells(i, 7).Value
                
                ' Print into Summary Table
                Ws.Range("I" & Summary_Table_Row).Value = Ticker_Name
                Ws.Range("J" & Summary_Table_Row).Value = Yearly_Change
                
                ' Format Colors
                If (Yearly_Change > 0) Then
                    Ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                ElseIf (Yearly_Change <= 0) Then
                    Ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                End If
                
                ' Print into Summary Table
                Ws.Range("K" & Summary_Table_Row).Value = (CStr(Percent_Change) & "%") 'CStr - convert a value to a string
                Ws.Range("L" & Summary_Table_Row).Value = Total_Ticker_Volume
                
                ' Add 1 to the summary table row count
                Summary_Table_Row = Summary_Table_Row + 1
                
                ' Reset
                Yearly_Change = 0
                Close_Price = 0
                Open_Price = Ws.Cells(i + 1, 3).Value
                Total_Ticker_Volume = 0
                
            Else
                ' Sum Total Ticker Volume
                Total_Ticker_Volume = Total_Ticker_Volume + Ws.Cells(i, 7).Value
                
            End If
                        
        Next i
        
     Next Ws
     
End Sub

Sub Challenge()

    ' Set Ws as a worksheet.
    Dim Ws As Worksheet
      
        ' Loop Worksheets.
        For Each Ws In Worksheets
        
            ' Variables and Values.
            Dim Greatest_Increase As Double
            Dim Greatest_Decrease As Double
            Dim Greatest_Volume As Double
            Dim r As Double
            
            Greatest_Increase = Ws.Cells(2, 11)
            Lastrow = Ws.Cells(Rows.Count, 1).End(xlUp).Row
            
            ' Headers
            Ws.Range("P1").Value = "Ticker"
            Ws.Range("Q1").Value = "Value"
            Ws.Range("O2").Value = "Greatest % Increase"
            Ws.Range("O3").Value = "Greatest % Decrease"
            Ws.Range("O4").Value = "Greatest Total Volume"
            
            'Loop
            For r = 2 To Lastrow
            
                'Find Values
                If Greatest_Increase < Ws.Cells(r, 11) Then
                    Greatest_Increase = Ws.Cells(r, 11)
                    End If
                If Greatest_Decrease > Ws.Cells(r, 11) Then
                    Greatest_Decrease = Ws.Cells(r, 11)
                    End If
                If Greatest_Volume < Ws.Cells(r, 12) Then
                    Greatest_Volume = Ws.Cells(r, 12)
                    End If
             
             Next r
            
            ' Print Values
            Ws.Range("Q2") = Greatest_Increase
            Ws.Range("Q3") = Greatest_Decrease
            Ws.Range("Q4") = Greatest_Volume
        
        Next Ws
        
End Sub


