' create nested for-loop that runs through all stocks for one year and outputs:
' - ticker symbol (stock name)
' - yearly change (closing price at end of year - opening price at start of year)
' - percent change ((closing price - opening price)/opening price)
'       - conditonal formatting to fill (+)change in green and (-)change in red
' - total stock volume
'----------------------------------------------------------------------------------------------------------

Sub stock_calculator()

'-- SET-UP LOOP THROUGH ALL SHEETS ----------------------------

'establish how many sheets are in the workbook
Dim sheetnumber As Integer
sheetnumber = ActiveWorkbook.Worksheets.Count

'tell it to loop through all the sheets
For k = 1 To sheetnumber
    'just in case, set the iterations sheet to active
    ActiveWorkbook.Worksheets(k).Activate

    '------ SET-UP SUMMARY TABLE HEADERS WITHIN CURRENT SHEET ------------------------

    'set location for summary table (I through L), print column titles in first row
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"

    '------ SET-UP VARIABLES FOR $CHANGE, %CHANGE, AND TOTAL VOLUME CALCULATIONS --------------------

    'Declare variable names and types
    Dim open_price As Double
    Dim close_price As Double
    Dim price_change As Double
    Dim percent_change As Double
    Dim T_stock_volume As Double

    'set opening price to the first price in the <open> column
    '   -value of opening price will reset with each loop
    open_price = Cells(2, 3).Value

    'set all other variables' initial value to 0
    close_price = 0
    price_change = 0
    percent_change = 0
    T_stock_volume = 0

    '------ CREATE LOOP WITHIN CURRENT SHEET --------------------

    'have to do some set-up before starting the loop:

    'set the count for the number of rows in the full worksheet as a varible
    Dim lastrow_sheet As Long
    lastrow_sheet = ActiveSheet.UsedRange.Rows.Count

    'set variable to hold last row value of summary table
    '   - since the blank table would start at row 2, set this initial variable value = 2
    '   - value of last row will increase with each loop
    Dim lastrow_summary As Integer
    lastrow_summary = 2

    'Now start the loop:

    'have the loop go through row 2 to lastrow_sheet
    For i = 2 To lastrow_sheet

        'If the value in (row i+1, <ticker> column) is different from the value before it (row i, <ticker>)...
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

            'Then start filling in the summary table row for the stock in (row i, <ticker>):

            'print value in <ticker> into the Ticker column
            Cells(lastrow_summary, 9).Value = Cells(i, 1).Value

            'set value in <close> as close_price for that stock
            close_price = Cells(i, 6).Value

            'retrieve open_price value and calulate price_change
            price_change = close_price - open_price

            'print price_change value into Yearly Change column
            Cells(lastrow_summary, 10).Value = price_change

            'check if it's a (+) or (-) change and fill Yearly Change cell with Green or Red accordingly
            If (price_change > 0) Then
                Cells(lastrow_summary, 10).Interior.ColorIndex = 4
            ElseIf (price_change <= 0) Then
                Cells(lastrow_summary, 10).Interior.ColorIndex = 3
            End If

            'check that open_price is not 0, then calulate percent_change
            If open_price <> 0 Then
                percent_change = (price_change / open_price) * 100
            End If

            'print percent_change value into Percent Change column with % format
            Cells(lastrow_summary, 11).Value = (CStr(percent_change) & "%")

            'add value in <vol> to T_stock_volume (will add values from other rows of that stock through the ElseIf condition for the main i loop)
            T_stock_volume = T_stock_volume + Cells(i, 7).Value

            'print T_stock_volume value into Total Stock Volume column
            Cells(lastrow_summary, 12).Value = T_stock_volume

            'Now reset values for next loop:

            'add 1 to lastrow_summary to reset the row count for the next loop
            lastrow_summary = lastrow_summary + 1

            'reset close_price and price_change to 0 since we'll be working with a new stock
            close_price = 0
            price_change = 0

            'find and set open_price for next stock so it can be used in the caculations in the next loop
            open_price = Cells(i + 1, 3).Value
            
            'reset T_stock_volume to 0 since we'll be working with a new stock
            T_stock_volume = 0
    
        Else
        T_stock_volume = T_stock_volume + Cells(i, 7).Value
        'this Else condition will add all the values in <vol> for the same stock listed before the last <vol> value into T_stock_volume
        'value for T_stock_value will be added to and kept until the stock's last row is reached through the initial If-Then condition
        'If-Then condition will then add last <vol> value for the stock and print total stock volume into the summary table

        End If

    Next i

    '------ SET-UP BONUS SUMMARY TABLE HEADERS WITHIN CURRENT SHEET ---------------------
   
    'set column titles in first row of column P and Q
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
    
    'set row titles in row 2, 3, and 4 of column O
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"

    'declare variables for max % increase, max % decrease, and max total volume
    'define variable values as equal to 0
    Dim max_percent_inc As Double
    max_percent_inc = 0
    Dim max_percent_dec As Double
    max_percent_dec = 0
    Dim max_T_volume As Double
    max_T_volume = 0
    
    'declare variables to hold the ticker names for max % increase, max % decrease, and max total volume
    'define variables as place holders (= " ")
    Dim GPI_ticker_name As String
    GPI_ticker_name = " "
    Dim GPD_ticker_name As String
    GPD_ticker_name = " "
    Dim GTV_ticker_name As String
    GTV_ticker_name = " "
    
    'create loop to run through row 2 until the lastrow of the summary table (defined in previous loop)
    For i = 2 To lastrow_summary
    
        'If the value under Percent Change is greater than value of the variable max_percent_inc, Then...
        If (Cells(i, 11).Value > max_percent_inc) Then
            'change the value of max_percent_inc to the value under Percent Change
            max_percent_inc = Cells(i, 11).Value
            'and set the ticker name for GPI_ticker_name to the value listed under the Ticker column for that row
            GPI_ticker_name = Cells(i, 9).Value
            
        'Otherwise, if the value under Percent change is less than the value of the variable max_percent_dec, Then...
        ElseIf (Cells(i, 11).Value < max_percent_dec) Then
            'change the value of max_percent_dec to the value under Percent Change
            max_percent_dec = Cells(i, 11).Value
            'and set the ticker name for GPD_ticker_name to the value listed under the Ticker column for that row
            GPD_ticker_name = Cells(i, 9).Value
            
        End If
        
        'If the value under Total Stock Volume is greater than the value of the variable max_T_volume, Then...
        If (Cells(i, 12).Value > max_T_volume) Then
            'change the value of max_T_volume to the value under Total Stock Volume
            max_T_volume = Cells(i, 12).Value
            'and set the ticker name for GTV_ticker_name to the value listed under the Ticker column for that row
            GTV_ticker_name = Cells(i, 9).Value
            
        End If
        
    Next i
    
    'print the values and ticker names for GPI, GPD, and GTV into their respective cells in the new summary table
    Range("Q2").Value = max_percent_inc
    Range("P2").Value = GPI_ticker_name
    Range("Q3").Value = max_percent_dec
    Range("P3").Value = GPD_ticker_name
    Range("Q4").Value = max_T_volume
    Range("P4").Value = GTV_ticker_name
    

Next k

End Sub
