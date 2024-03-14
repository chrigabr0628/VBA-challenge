An overview of the analysis: 

The purpose of the analysis was to identify the changes of stock volume within three years of data.


The results: 

Based on the data, ticker 'RYU' had the greatest increase in percentage change, ticker 'RKS' had the greatest decrease in percentage change, and ticker 'ZQD' had the greatest total volume as well. 


Methodology: 

Firstly, I created a 'for loop' to loop through all three sheets of data. I created variables to retrieve the ticker symbol, yearly change, percentage change, total stock volume, greatest percentage increase, greatest percentage decrease, and greatest total volume. From there, I filled each column with its respective values and was able to determine the most and least successful years for the stock market based on the data.


Conclusion:

Based on the data, year 2019 had the greatest increase in percentage change, in addition to the greatest decrease in percentage change and the greatest total volume as well. Despite having the greatest decrease in  percentage change, 2019 appears to have been a great year for the stock market based on these findings. 2018 appears to have been the worst year based on these findings as well.


Next Steps: 

Fix the dates and remove irrelevant data such as 'high' and 'low'.






Code Source by Tutor:
    
    Dim volume As LongLong
   
    
    Dim info As LongLong
  

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





Code Source by TA:

        Else
          
              volume = volume + ws.Cells(i, 7).Value
