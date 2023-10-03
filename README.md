# VBA-Challenge
Module 2
A Majority of the code lines came from adapted and learned code used during the Module 2 week course material.

The Modules Name / Task
Module 1 : Columns- Is used to create and name the first 4 columns of the Challenge
Module 2: Consolidated ticker - Consolidates all of the different ticker names into a single category and adds their associated stock volume
Module 3: Yearly change - Solves for the yearly change of each ticker and adjust the color of the cell depending if it was a positive or negative change
Module 4: Percentage Change - Solves for the percentage change of each of the consolidated tickers
Module 5: Create Greatest Columns - Like module 1 is setting up 2 new columns and 3 rows for information to be placed
Module 6: Solve for Greatest - Is searching the worksheets for the greatest increase / decrease and total volume per worksheet

Specific code used from others include the following

1. In Module 6 (Solve_for_Greatest) section the idea to use the  m = Application.WorksheetFunction.Max(r) and .Min functions comes from the website // https://www.educba.com/vba-max/

2. In Module 3 (Yearly_change) most of the code for the changing of colors came from the classwork material however the Code:  If ws.Cells(i, 10) <> "" Then to remove colors being filled into blank spaces came from the website // https://stackoverflow.com/questions/13705663/excel-user-defined-function-change-the-cells-color

3. In Module 4 (Percentage_change) the code to change the format of the percentages : ws.Cells(consolidated_info_row, 11).NumberFormat = "0.00%" came from the website // https://www.excelfunctions.net/vba-formatpercent-function.html

4.Both Module 1 (Columns) and Module 5 (Create_Greatest_Columns) both used a "UBound(name) function which I found from the website // https://stackoverflow.com/questions/74148145/get-the-desired-name-split-in-vba
