# A simple VBA script
Module 2 challenge - UFT Data Analytics Bootcamp
1. The VBA script is in .bas format. It contains two sub-routines. The main sub-routine named worksheet_looper() declares a variable of type Worksheet and 'for' loops through each worksheet in the workbook. At each iteration, the current worksheet is activated, its name is printed on the console and another sub-routine named stock_analysis() is executed on it.
2. In the stock-analysis worksheet, first the total number of rows (excluding the very first row) is counted and the last row is identified using built-in function.
3. Then, the columns outside of the pre-existing data in the worksheet are cleared off of their contents and formats. The same columns are then given a fixed column width using the ColumnWidth attribute
4. The values that are to be computed are declared as variables as their respective datatypes.
5. The result columns are hardcoded.
6. A 'tickerCount' is initialized to keep track of how many 'groups' of tickers have been found
7. An outer loop is created to iterate through all the rows of data in the current worksheet
    (1) The current ticker label and a ticker counter for it are initialized 
    (2) An inner conditional loop is initialized from here. 
            (A) It iterates until the next row is NOT same as the current one. 
            (B) The index at the end is the end point of a ticker group.
    (3) The range from the outer loop index to the inner loop index allows us to compute the yearly_change, percent_change and the volume for a specific ticker.
    (4) The results are recorded in the hardcoded columns where the rows are identified by the tickerCount variable
    (5) Conditional formatting is applied for the yearly_change and percent_change columns
    (6) Lastly, the index for the outer loop is updated with the updated index of the inner loop. As a result the outer loop is able to iterate really fast
8. The built-in max and min functions are used to find the max and min volumes and percent changes down the resulting columns, and the corresponding row indices are found using the built-in match function
9. The final results are retrieved from the matched rows and recorded into hardcoded column locations and formatted as desired.

The script is able to generate the results for the whole workbook in 5 to 8 seconds in a Win 10, intel i7, 32 GB RAM system.
