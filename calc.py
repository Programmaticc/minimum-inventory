import pandas as pd
import math
import os
import glob

# Function to calculate minimum stock
def calculate_min_stock(z_value, avg_lead_time_days, daily_demand_mean, daily_demand_std_dev, daily_lead_time_std_dev):
    min_stock = z_value * math.sqrt((avg_lead_time_days * daily_demand_std_dev**2) + (daily_demand_mean**2 * daily_lead_time_std_dev**2))
    return round(min_stock, 2)

# Input of values
avg_lead_time_days = float(input("Please enter the average delivery time (in days): "))
daily_lead_time_std_dev = avg_lead_time_days * 0.3  # Standard deviation of delivery time as 30% of the average delivery time
z_value = 1.65  # Z-value for 95% service level

# Folder path
folder_path = r"C:\Users\Desktop"

# Search for all Excel files in the specified folder
file_paths = glob.glob(os.path.join(folder_path, '*_output.xlsx'))

for file_path in file_paths:
    # Read data from the Excel file
    data = pd.read_excel(file_path, header=None)

    # Aggregation of demand data per item and depot over the years
    grouped_data = data.groupby([0, 1])  # Group by item and depot

    # Save results
    results = []

    # Loop over each group
    for (item, depot), group in grouped_data:
        demand_data = group.iloc[:, 3:14].values.flatten()  # Summarize demand data for all years
        non_zero_demand_data = demand_data[demand_data != 0]  # Only non-zero sales data

        if len(non_zero_demand_data) == 0:
            continue  # If no sales are available, skip this group

        monthly_demand_mean = non_zero_demand_data.mean()  # Average monthly demand
        monthly_demand_std_dev = non_zero_demand_data.std()  # Standard deviation of monthly demand

        daily_demand_mean = monthly_demand_mean / 20  # Conversion of monthly demand to daily demand
        daily_demand_std_dev = monthly_demand_std_dev / math.sqrt(20)  # Conversion of monthly standard deviation to daily standard deviation

        # Calculation of minimum stock
        monthly_min_stock = calculate_min_stock(z_value, avg_lead_time_days, daily_demand_mean, daily_demand_std_dev, daily_lead_time_std_dev)

        # Save results
        results.append([
            item, 
            depot, 
            monthly_min_stock, 
            round(monthly_demand_std_dev, 2),
            round(monthly_demand_mean, 2),
        ])

    # Write results to a new Excel file
    results_df = pd.DataFrame(results, columns=[
        'Item', 
        'Depot', 
        'Monthly Minimum Stock', 
        'Monthly Standard Deviation of Demand', 
        'Monthly Mean of Sales'
    ])
    
    # Create the new filename
    result_file_path = file_path.replace('_output.xlsx', '_result.xlsx')
    
    # Save the results
    results_df.to_excel(result_file_path, index=False)

print("Calculations completed and results saved.")
