import pandas as pd
import datetime
import openpyxl
from openpyxl.styles import NamedStyle

def get_filename(prefix, file_type):
    today = datetime.date.today()

    # Check if today is Monday
    if today.weekday() == 0:  # Monday has a weekday of 0
        target_date = today - datetime.timedelta(days=3)  # Subtract 3 days to get Friday
    else:
        target_date = today - datetime.timedelta(days=1)  # Subtract 1 day to get the previous day

    # Construct the filename
    filename = f"{prefix}_{file_type}_{target_date.strftime('%Y%m%d')}.csv"
    return filename

def process_dataframe(df, columns_to_drop, columns_to_sum, account_column_name):
    df = df.drop(columns=columns_to_drop, errors='ignore')
    df = df.sort_values(by=account_column_name)
    df = df[df[account_column_name].isin(accounts_to_keep)]
    df = df.groupby(account_column_name)[columns_to_sum].sum().reset_index()

    # Compute the total row
    summed_values = df[columns_to_sum].sum()
    summed_values[account_column_name] = 'Total'
    df = pd.concat([df, pd.DataFrame([summed_values])], ignore_index=True)

    return df

def save_to_excel(df, filename):
    # Define style
    number_format_style = NamedStyle(name="number_format_style", number_format="#,##0.00_);[Red](#,##0.00)")
    
    df.to_excel(filename, index=False, engine='openpyxl')
    
    wb = openpyxl.load_workbook(filename)
    ws = wb.active

    # Apply formatting
    for column in ws.columns:
        max_length = 0
        column = [cell for cell in column]
        for cell in column:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(cell.value)
            except:
                pass
        adjusted_width = (max_length + 2) if max_length < 20 else 20
        ws.column_dimensions[column[0].column_letter].width = adjusted_width

        for cell in column:
            if isinstance(cell.value, (int, float)):
                cell.style = number_format_style

    wb.save(filename)

# List of account numbers to keep
accounts_to_keep = [
    "66EG99OL", "66EG99WY", "66TX99JP", "66TX99RC", "66TX99DJ", "66TX99OL", 
    "66TX99JK", "66TX99CB", "66TX99MF", "66TX99OC", "66TX99DS", "66EG99E1", 
    "66EG99EA", "66EG99EG", "66TX99AP", "66TX99CC", "66TX99CP", "66TX99JD", 
    "66TX99JR", "66TX99KS", "66TX99OE", "66TX99OG", "66TX99OX", "66TX99ER", 
    "66TX99FI", "66TX99VK", "EGTXMUNI", "House", "66TX99TR", "66TX99WY"
]
columns_to_sum_daily = [
     "TodayRealizedPL", "TodayCouponInterest",
     "TodayTotalPL",
]
columns_to_drop_daily = [
    "YTDUnrealizedPL", "YTDRealizedPL", "YTDCouponInterest",
    "YTDFundingInterest", "YTDTotalPL", "ClosingPrice", 
    "factor", "Rate", "TradeQuantity", "SettleQuantity", 
    "TradeMarketValue", "SettleMarketValue",
    "OfficeRR", "Currency", "SecurityType", "Symbol", 
    "CUSIP", "Sedol", "SecurityDescription", "MTDFundingInterest",
    "TodayFundingInterest", "TodayUnrealizedPL", "MTDUnrealizedPL",
    "MTDRealizedPL", "MTDCouponInterest", "MTDTotalPL"
]
columns_to_drop_realized = [
    "MTDRealizedPL", "YTDRealizedPL", "Today_FundingInterest", 
    "MTDFundingInterest", "YTDFundingInterest", "MTDCouponInterest",
    "YTDCouponInterest", "MTDCouponPayment", "YTDCouponPayment",
    "TodayPrincipalPaydown", "MTDPrincipalPaydown", "YTDPrincipalPaydown",
    "CUSIP"
]
columns_to_sum_realized = ["TodayRealizedPL", "TodayCouponInterest", "TodayCouponPayment"]

# RealizedPL Data
tx_realized = pd.read_csv(get_filename("TX", "RealizedPL"))
eg_realized = pd.read_csv(get_filename("EG", "RealizedPL"))
merged_realized = pd.concat([tx_realized, eg_realized], ignore_index=True)
processed_realized = process_dataframe(
    merged_realized,
    columns_to_drop_realized, # Add other columns here
    columns_to_sum_realized,
    account_column_name="Account"
)
processed_realized.columns = [col if col in ["Account"] else "Realized PL - " + col for col in processed_realized.columns]
save_to_excel(processed_realized, 'Formatted_Merged_RealizedPL.xlsx')

# DailyPL Data
tx_daily = pd.read_csv(get_filename("TX", "DailyPL"))
eg_daily = pd.read_csv(get_filename("EG", "DailyPL"))
merged_daily = pd.concat([tx_daily, eg_daily], ignore_index=True)
processed_daily = process_dataframe(
    merged_daily,
    columns_to_drop_daily,
    columns_to_sum_daily, # Add other columns here
    account_column_name="AccountNumber"
)
processed_daily.columns = [col if col in ["AccountNumber"] else "Daily PL - " + col for col in processed_daily.columns]
save_to_excel(processed_daily, 'Formatted_Merged_DailyPL.xlsx')

print('Files processed and Excel formatting applied successfully!')

# Step 1: Merge the dataframes on the account columns
comparison_df = pd.merge(
    processed_realized[['Account', 'Realized PL - TodayRealizedPL', 'Realized PL - TodayCouponInterest']],
    processed_daily[['AccountNumber', 'Daily PL - TodayRealizedPL', 'Daily PL - TodayCouponInterest']],
    left_on='Account',
    right_on='AccountNumber',
    how='outer'
)

# Step 2: Calculate differences
comparison_df['Difference - TodayRealizedPL'] = comparison_df['Realized PL - TodayRealizedPL'] - comparison_df['Daily PL - TodayRealizedPL']
comparison_df['Difference - TodayCouponInterest'] = comparison_df['Realized PL - TodayCouponInterest'] - comparison_df['Daily PL - TodayCouponInterest']

# Step 3: Create the final output DataFrame in the desired order and add an empty column
final_output = comparison_df[[
    'Account',
    'Realized PL - TodayRealizedPL',
    'Daily PL - TodayRealizedPL',
    'Difference - TodayRealizedPL',
    'Realized PL - TodayCouponInterest',
    'Daily PL - TodayCouponInterest',
    'Difference - TodayCouponInterest'
]]
final_output.insert(4, '', '')  # Inserting an empty column

# Save the final output to an Excel file
save_to_excel(final_output, 'Comparison_File.xlsx')

print('Comparison file created successfully!')