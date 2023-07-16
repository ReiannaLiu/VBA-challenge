# VBA-challenge

This challenge is intended to use VBA scripting to analyze generated stock market data.

For each ticker, we calculate the total stock volume of the stock using aggregate function Sum to sum over all the stock volume for the year.

Then the yearly change from the opening price at the beginning of a given year to the closing price at the end of that year is calculated as the closing price at the last row of the ticker minus the opening price at the first row of the ticker. We use TickerStart to keep track of the first row of the ticker, and the current row index i to keep track the last row of the ticker.

The percentage change from the opening price at the beginning of a given year to the closing price at the end of that year is calculated as the yearly change divided by the opening price at the beginning of a given year.

In this challenge, we also summarize the greatest % increase, greatest % decrease, and greatest total stock volume with corresponding ticker.
