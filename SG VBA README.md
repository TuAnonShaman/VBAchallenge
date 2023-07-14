# VBA README - Stephen Grantham

This challenge required VBA code to quickly summarize sheets of daily stock data from an entire year.  


The summary table included columns for...

    Ticker - grabbed ticket (from column A) in 'for loop' whenever column A changed value in the next row
    
    Yearly Change - grabbed opening price (from column C) in 'for loop' whenever column A changed value from previous row.  Closing price (from column F) when value changed in next row.
        Yearly Change = Close Price - Open Price
        
    Percent Change - (Close Price - Open Price) / (Open Price)
        
        Conditionally formatted to show green for positive, red for negative, and nothing for 0.  

    Total Stock Volume - calculated by summing a 'TotalStockVolume' (from column G) for a ticker until column A changed


An additional summary table included values for...

    Greatest % Increase - looped 'Percent Change' column to find highest value
    
    Greatest % Decrease - looped 'Percent Change' column to find the lowest value
    
    Greatest Total Volume - looped 'Total Stock Volume' column to find highest value
    
Finally, the code was modified to run for worksheets of all three years from 2018 - 2020
