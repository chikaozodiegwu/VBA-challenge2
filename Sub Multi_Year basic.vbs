Sub Multi_Year()

Dim CurrentWs As Worksheet
    Dim Summary_Table_Header As Boolean
    Dim COMMAND_SPREADSHEET As Boolean
    
    Summary_Table_Header = False
    COMMAND_SPREADSHEET = True
    
    For Each CurrentWs In Worksheets
    
        
        Dim Ticker_Name As String
        Ticker_Name = " "
        
        
        Dim Ticker_Vol As Double
        Ticker_Vol = 0
        
        
        Dim Open_Price As Double
        Open_Price = 0
        Dim Close_Price As Double
        Close_Price = 0
        Dim Yearly_Change As Double
        Yearly_Change = 0
        Dim Percent_Change As Double
        Percent_Change = 0
        
        Dim Summary_Table_Row As Long
        Summary_Table_Row = 2

        Dim Lastrow As Long
        Dim i As Long
        
        Lastrow = CurrentWs.Cells(Rows.Count, 1).End(xlUp).Row

         
        If Summary_Table_Header Then
           
            CurrentWs.Range("H1").Value = "Ticker"
            CurrentWs.Range("I1").Value = "Yearly Change"
            CurrentWs.Range("J1").Value = "Percent Change"
            CurrentWs.Range("K1").Value = "Total Stock Volume"

        Else
            
            Summary_Table_Header = True
        End If
        
        Open_Price = CurrentWs.Cells(2, 3).Value
        
        
        For i = 2 To Lastrow
        
      
            
            If CurrentWs.Cells(i + 1, 1).Value <> CurrentWs.Cells(i, 1).Value Then
            
                
                Ticker_Name = CurrentWs.Cells(i, 1).Value
                
               
                Close_Price = CurrentWs.Cells(i, 6).Value
                Yearly_Change = Close_Price - Open_Price
                
                If Open_Price <> 0 Then
                    Percent_Change = (Yearly_Change / Open_Price) * 100
                Else
                   
                    MsgBox ("For " & Ticker_Name & ", Row " & CStr(i) & ": Open Price =" & Open_Price & ". Fix <open> field manually and save the spreadsheet.")
                End If
                
               
                Ticker_Vol = Ticker_Vol + CurrentWs.Cells(i, 7).Value
              
                
                
                CurrentWs.Range("H" & Summary_Table_Row).Value = Ticker_Name
               
                CurrentWs.Range("I" & Summary_Table_Row).Value = Yearly_Change
                
                If (Yearly_Change > 0) Then
                    
                    CurrentWs.Range("I" & Summary_Table_Row).Interior.ColorIndex = 4
                ElseIf (Yearly_Change <= 0) Then
                    
                    CurrentWs.Range("I" & Summary_Table_Row).Interior.ColorIndex = 3
                End If
                
                 
                CurrentWs.Range("J" & Summary_Table_Row).Value = (CStr(Percent_Change) & "%")
                
                CurrentWs.Range("K" & Summary_Table_Row).Value = Ticker_Vol
                

                Summary_Table_Row = Summary_Table_Row + 1
                
                Yearly_Change = 0
               
                Close_Price = 0
                
                Open_Price = CurrentWs.Cells(i + 1, 3).Value
              
                Percent_Change = 0
                Ticker_Vol = 0

            Else
                Ticker_Vol = Ticker_Vol + CurrentWs.Cells(i, 7).Value
            
            End If

        Next i
    
    Next CurrentWs

End Sub
