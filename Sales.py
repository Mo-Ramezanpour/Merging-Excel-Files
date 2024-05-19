import pandas as pd
import datetime as dt
import calendar

# List of Excel file paths to append
excel_files = ["Ireland.xlsx", "Japan.xlsx", "Norway.xlsx"]

# Initialize an empty list to store DataFrames
dfs = []

# Iterate over each Excel file and read data into a DataFrame
for file in excel_files:
    df = pd.read_excel(file)
    df['Month'] = pd.DatetimeIndex(df['Date']).month
    df["Year"] = pd.DatetimeIndex(df["Date"]).year
    df['Sales'] = df['Quantity'] * df['Price']
    df['Sales'] = df['Sales'].map('${:,.0f}'.format)
    df['Date'] = df['Date'].dt.strftime('%m/%d/%Y')
    df["Month_Name"] = df["Month"].apply(lambda x: calendar.month_abbr[x])
    dfs.append(df)

# Concatenate all DataFrames into a single DataFrame
combined_data = pd.concat(dfs, ignore_index=True)

# Convert 'Sales' column back to float for calculations
combined_data['Sales'] = combined_data['Sales'].replace(
    '[\\$,]', '', regex=True).astype(float)

# Calculate the total sales for each month
monthly_sales = combined_data.groupby(['Year', 'Month'])[
    'Sales'].transform('sum')

# Add a new column for the total monthly sales
combined_data['Monthly_Sales'] = monthly_sales

# Calculate the total sales for each year
yearly_sales = combined_data.groupby(['Year'])['Sales'].transform('sum')

# Calculate the percentage of total sales for each day and each month
combined_data['%_of_Total_Sales'] = combined_data['Sales'] / monthly_sales
combined_data['%_of_Total_Sales_in_Year'] = monthly_sales / yearly_sales


# Format the percentage and monthly sales to be more readable
combined_data['%_of_Total_Sales'] = combined_data['%_of_Total_Sales'].map(
    '{:.2%}'.format)
combined_data['%_of_Total_Sales_in_Year'] = combined_data['%_of_Total_Sales_in_Year'].map(
    '{:.2%}'.format)
combined_data['Monthly_Sales'] = combined_data['Monthly_Sales'].map(
    '${:,.0f}'.format)

# Convert 'Sales' column back to currency format
combined_data['Sales'] = combined_data['Sales'].map('${:,.0f}'.format)

# Specify the path for the output combined Excel file
output_file = "Total_Sales.xlsx"

# Write the combined DataFrame to an Excel file
combined_data.to_excel(output_file, index=False)
