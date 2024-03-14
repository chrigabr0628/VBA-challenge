Attribute VB_Name = "Module1"
Sub VBAChallenge()

    Dim ws As Worksheet
    
    For Each ws In ThisWorkbook.Worksheets
    

    Dim symbol As String
    
    Dim yearly As Double
    yearly = 0
    
    Dim percent As Double
    percent = 0
    
    Dim volume As LongLong
    volume = 0
    
    Dim info As LongLong
    info = 2
  
    Dim openprice As Double
    
    openprice = ws.Cells(2, 3).Value
    
    ws.Cells(1, 10).Value = "ticker"
    ws.Cells(1, 11).Value = "yearly change"
    ws.Cells(1, 12).Value = "percentage change"
    ws.Cells(1, 13).Value = "total stock volume"
    
    Dim LastRow As Long
    
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
  
    
    
    For i = 2 To LastRow

 
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
     
              symbol = ws.Cells(i, 1).Value
              
              yearly = ws.Cells(i, 6).Value - openprice
              
              percent = yearly / openprice
              
              volume = volume + ws.Cells(i, 7).Value
            
            
            
              ws.Range("J" & info).Value = symbol
              
              ws.Range("K" & info).Value = yearly
              
              ws.Range("L" & info).Value = percent
              
              ws.Range("M" & info).Value = volume
              
              
              info = info + 1
             
              volume = 0
    
              openprice = ws.Cells(i + 1, 3).Value
              
          Else
          
              volume = volume + ws.Cells(i, 7).Value
          
          
          End If
          
        Next i
     
     
     
          
         
                    Dim increase As Long
                    increase = 0
                    
                    Dim decrease As Long
                    percent = 0
                    
                    Dim total As Long
                    total = 0
                
                
                    ws.Cells(2, 16).Value = "Greatest % increase"
                    ws.Cells(3, 16).Value = "Greatest % decrease"
                    ws.Cells(4, 16).Value = "Greatest total volume"
                    ws.Cells(1, 17).Value = "ticker"
                    ws.Cells(1, 18).Value = "value"
                    
    
     
      Next ws
  
        
 

    

End Sub

