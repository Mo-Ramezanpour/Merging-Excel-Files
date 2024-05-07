# Merging-Excel-Files

This Python code is designed to combine sales data from multiple Excel files, each representing sales data from different countries (Ireland, Japan, Norway). The code utilizes the pandas library to read and manipulate the data. It performs various operations including:

1. Reading data from each Excel file into a DataFrame.
2. Adding columns for month, year, total sales, and month name.
3. Concatenating all DataFrames into a single DataFrame.
4. Converting sales data to a float format for calculations.
5. Calculating total sales for each month and year.
6. Calculating the percentage of total sales for each day and each month, as well as the percentage of total sales in the year.
7. Adding a column for total monthly sales.
8. Formatting percentage and sales values for readability.
9. Writing the combined DataFrame to a new Excel file named "Total_Sales.xlsx".

The resulting Excel file will contain the combined sales data along with additional columns for total monthly sales, percentage of total sales, and percentage of total sales in the year for each entry.
