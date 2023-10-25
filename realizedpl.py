import pandas as pd
import datetime
import openpyxl
from openpyxl.styles import NamedStyle

# Function to determine the filename for the previous day (or Friday if today is Monday)
def get_previous_day_filename(prefix):
    today = datetime.date.today()
    
    # Check if today is Monday
    if today.weekday() == 0:  # Monday has a weekday of 0
        target_date = today - datetime.timedelta(days=3)  # Subtract 3 days to get Friday
    else:
        target_date = today - datetime.timedelta(days=1)  # Subtract 1 day to get the previous day

    # Construct the filename
    filename = f"{prefix}_RealizedPL_{target_date.strftime('%Y%m%d')}.csv"
    return filename

# Determine the filenames for TX and EG
tx_filename = get_previous_day_filename("TX")
eg_filename = get_previous_day_filename("EG")

# Load the two CSV files into dataframes
tx_df = pd.read_csv(tx_filename)
eg_df = pd.read_csv(eg_filename)

# Concatenate the two dataframes vertically
merged_df = pd.concat([tx_df, eg_df], ignore_index=True)

# List of columns to drop
columns_to_drop = [
    "MTDRealizedPL", "YTDRealizedPL", "Today_FundingInterest", 
    "MTDFundingInterest", "YTDFundingInterest", "MTDCouponInterest",
    "YTDCouponInterest", "MTDCouponPayment", "YTDCouponPayment",
    "TodayPrincipalPaydown", "MTDPrincipalPaydown", "YTDPrincipalPaydown",
    "CUSIP"
]

# Drop the columns
merged_df = merged_df.drop(columns=columns_to_drop, errors='ignore')

merged_df = merged_df.sort_values(by="Account")

# List of account numbers to keep
accounts_to_keep = [
    "66EG99OL", "66EG99WY", "66TX99JP", "66TX99RC", "66TX99DJ", "66TX99OL", 
    "66TX99JK", "66TX99CB", "66TX99MF", "66TX99OC", "66TX99DS", "66EG99E1", 
    "66EG99EA", "66EG99EG", "66TX99AP", "66TX99CC", "66TX99CP", "66TX99JD", 
    "66TX99JR", "66TX99KS", "66TX99OE", "66TX99OG", "66TX99OX", "66TX99ER", 
    "66TX99FI", "66TX99VK", "EGTXMUNI", "House", "66TX99TR", "66TX99WY"
]

# Filter the dataframe to retain only the specified account numbers
merged_df = merged_df[merged_df['Account'].isin(accounts_to_keep)]

# Columns to sum for each account number
columns_to_sum = ["TodayRealizedPL", "TodayCouponInterest", "TodayCouponPayment"]

# Group by AccountNumber and sum the specified columns
merged_df = merged_df.groupby('Account')[columns_to_sum].sum().reset_index()

# Compute the sum for each of the columns_to_sum and append to the end
summed_values = merged_df[columns_to_sum].sum()
summed_values['Account'] = 'Total'
merged_df = pd.concat([merged_df, pd.DataFrame([summed_values])], ignore_index=True)

# Define a style for the desired number format
number_format_style = NamedStyle(name="number_format_style", number_format="#,##0.00_);[Red](#,##0.00)")

# Save the dataframe to an Excel file
excel_file = 'Formatted_Merged_RealizedPL.xlsx'
merged_df.to_excel(excel_file, index=False, engine='openpyxl')
