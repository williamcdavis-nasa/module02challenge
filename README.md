Module 2 Challenge
Will Davis | william.c.davis@nasa.gov
-----
Retrieval of Data (20 points)
    The script loops through one quarter of stock data and reads/ stores all of the following values from each row:
    	ticker symbol (5 points)
    	volume of stock (5 points)
    	open price (5 points)
    	close price (5 points)

    See attached files in submittal.
	Excel file: Multiple_year_stock_data-wcd
	VBS file: quarter_single.vbs or quarter_all.vbs
    	ticker symbol
            Dim Ticker As String
    	volume of stock
            Dim Vol as Double
    	open price
            Dim OpenValue as Double
    	close price
            Dim CloseValue as Double
    	Screencaps: Q1_screenshot, Q2_screenshot, Q3_screenshot, Q4_screenshot

Column Creation (10 points)
   	On the same worksheet as the raw data, or on a new worksheet all columns were correctly created for:
        ticker symbol (2.5 points)
        total stock volume (2.5 points)
    	quarterly change ($) (2.5 points)
    	percent change (2.5 points)

    See attached files in submittal.
    	Excel file: Multiple_year_stock_data-wcd
    	VBS file: quarter_single.vbs or quarter_all.vbs
        	ticker symbol
                ws.Range("I" & SummaryRow).Value = "Ticker”
        	total stock volume
                ws.Range("J" & SummaryRow).Value = "Quarterly Change"
        	quarterly change
                ws.Range("J" & SummaryRow).Value = "Quarterly Change"
        	percent change
                ws.Range("K" & SummaryRow).Value = "Percent Change"
    	Screencaps: Q1_screenshot, Q2_screenshot, Q3_screenshot, Q4_screenshot


Conditional Formatting (20 points)
    Conditional formatting is applied correctly and appropriately to the quarterly change column (10 points)
    Conditional formatting is applied correctly and appropriately to the percent change column (10 points)

    See attached files in submittal.
    	Excel file: Multiple_year_stock_data-wcd
    	VBS file: quarter_single.vbs or quarter_all.vbs
        	quarterly change column  see code below
                Formatting
                    ws.Range("J" & SummaryRow).NumberFormat = "0.00"
                Conditional coloring
                    If QtrChangeNum > 0 Then
                        ws.Range("J" & SummaryRow).Interior.ColorIndex = 4
                    ElseIf QtrChangeNum < 0 Then
                        ws.Range("J" & SummaryRow).Interior.ColorIndex = 3
                    Else
                        ws.Range("J" & SummaryRow).Interior.ColorIndex = xlNone
                    End If
            percent change column
                	Formatting
                    ws.Range("J" & SummaryRow).NumberFormat = "0.00"	
    	Screencaps: Q1_screenshot, Q2_screenshot, Q3_screenshot, Q4_screenshot

Calculated Values (15 points)
    All three of the following values are calculated correctly and displayed in the output:
    	Greatest % Increase (5 points)
    	Greatest % Decrease (5 points)
    	Greatest Total Volume (5 points)

    See attached files in submittal.
        Excel file: Multiple_year_stock_data-wcd
        VBS file: quarter_single.vbs or quarter_all.vbs
            Greatest % Increase (5 points)  see code below
        	Greatest % Decrease (5 points)   see code below
            Greatest Total Volume (5 points)  see code below

                SumHigh = ws.Cells(2, 11)
                SumLow = ws.Cells(2, 11)
                SumVol = ws.Cells(2, 12)
    
                For j = jStart To (jEnd - 1)
                    If ws.Cells(j + 1, 11) > SumHigh Then
                        SumHigh = ws.Cells(j + 1, 11)
                        SumTickerHigh = ws.Cells(j + 1, 9)
                    ElseIf ws.Cells(j + 1, 11) < SumLow Then
                        SumLow = ws.Cells(j + 1, 11)
                        SumTickerLow = ws.Cells(j + 1, 9)
                    ElseIf ws.Cells(j + 1, 12) > SumVol Then
                        SumVol = ws.Cells(j + 1, 12)
                        SumTickerVol = ws.Cells(j + 1, 9)
                    End If
                Next j

        Screencaps: Q1_screenshot, Q2_screenshot, Q3_screenshot, Q4_screenshot

Looping Across Worksheet (20 points)
	The VBA script can run on all sheets successfully.

    See attached files in submittal.
    	Excel file: Multiple_year_stock_data-wcd
    	VBS file: quarter_all.vbs

GitHub/GitLab Submission (15 points)
    All three of the following are uploaded to GitHub/GitLab:
       	Screenshots of the results (5 points)
            See attached files in submittal.
              	Screencaps: Q1_screenshot, Q2_screenshot, Q3_screenshot, Q4_screenshot

        Separate VBA script files (5 points)
            See attached files in submittal.
            	VBS files: quarter_single.vbs or quarter_all.vbs

    	README file (5 points)
            See attached file in submital (this file)
            	ReadMe.md
