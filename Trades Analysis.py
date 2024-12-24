import pandas as pd
import matplotlib.pyplot as plt
from openpyxl import Workbook
from openpyxl.drawing.image import Image


# File: Data_Analyzed.py
def analyze_data():
    # Load your dataset
    data = pd.read_csv('data.csv', parse_dates=['date'])

    # Drop unwanted columns
    columns_to_drop = ['Unnamed: 0', 'Unnamed: 22', 'Unnamed: 23', 'Unnamed: 24', 'Unnamed: 25', 'Unnamed: 26']
    data = data.drop(columns=columns_to_drop)

    # Set initial balance
    balance = 100000

    # Initialize columns
    data['real_quantity'] = 0
    data['profitLose'] = 0
    data['commissions'] = 0
    data['neto'] = 0
    data['updated_balance'] = ''

    # Iterate through rows to calculate values
    for i in range(len(data)):
        # Calculate real_quantity
        data.loc[i, 'real_quantity'] = (balance * 0.1) / data.loc[i, 'buy_point']

        # Calculate profitLose
        if data.loc[i, 'action'] == 'BUY' and data.loc[i, 'pl'] == 'P':
            data.loc[i, 'profitLose'] = (data.loc[i, 'take_profit'] - data.loc[i, 'buy_point']) * data.loc[
                i, 'real_quantity']
        elif data.loc[i, 'action'] == 'BUY' and data.loc[i, 'pl'] == 'L':
            data.loc[i, 'profitLose'] = (data.loc[i, 'stop_loss'] - data.loc[i, 'buy_point']) * data.loc[
                i, 'real_quantity']
        elif data.loc[i, 'action'] == 'SELL' and data.loc[i, 'pl'] == 'P':
            data.loc[i, 'profitLose'] = (data.loc[i, 'buy_point'] - data.loc[i, 'take_profit']) * data.loc[
                i, 'real_quantity']
        elif data.loc[i, 'action'] == 'SELL' and data.loc[i, 'pl'] == 'L':
            data.loc[i, 'profitLose'] = (data.loc[i, 'buy_point'] - data.loc[i, 'stop_loss']) * data.loc[
                i, 'real_quantity']

        # Calculate commissions
        if data.loc[i, 'real_quantity'] > 250:
            data.loc[i, 'commissions'] = (((data.loc[i, 'real_quantity'] - 250) * 0.008) * 2) + 4
        else:
            data.loc[i, 'commissions'] = 4

        # Calculate neto
        data.loc[i, 'neto'] = data.loc[i, 'profitLose'] - data.loc[i, 'commissions']

        # Update balance for the next iteration
        balance += data.loc[i, 'neto']

        # Calculate updated_balance for each date
        if i == len(data) - 1 or data.loc[i, 'date'] != data.loc[i + 1, 'date']:
            data.loc[i, 'updated_balance'] = balance

    # Save the updated DataFrame to a new XLSX file with worksheet name 'Data'
    data.to_excel('data_analyzed.xlsx', sheet_name='Data', index=False)

    print("Data analysis completed and saved to 'data_analyzed.xlsx'")


# File: Summary.py
def create_summary():
    # Load the original dataset from the 'Data' worksheet of the XLSX file
    original_data = pd.read_excel('data_analyzed.xlsx', sheet_name='Data')
    original_data['date'] = pd.to_datetime(original_data['date'], format='%d/%m/%Y', errors='coerce')

    # Ensure 'date' is recognized as a datetime column
    original_data['Month'] = original_data['date'].dt.month

    # Create a summary DataFrame
    summary = pd.DataFrame({'Month': list(range(1, 13))})

    # Profits - count how many P i have in the column pl from the original data for each month
    profits_count = original_data[original_data['pl'] == 'P'].groupby('Month').size().reset_index(name='Profits')
    summary = pd.merge(summary, profits_count, how='left', on='Month')

    # Losses - count how many L i have in the column pl from the original data for each month
    losses_count = original_data[original_data['pl'] == 'L'].groupby('Month').size().reset_index(name='Losses')
    summary = pd.merge(summary, losses_count, how='left', on='Month')

    # Monthly Positions - sum of Profits and Losses for each month
    summary['Monthly Positions'] = summary['Profits'] + summary['Losses']

    # Hit Percentage - calculate the percentage of Profits out of Monthly Positions
    summary['Hit Percentage'] = (summary['Profits'] / summary['Monthly Positions']).fillna(0) * 100

    # Commission - sum of Commission for each month
    summary['Commission'] = original_data.groupby('Month')['commissions'].sum().reset_index()['commissions']

    # Neto - sum of Neto for each month
    summary['Neto'] = original_data.groupby('Month')['neto'].sum().reset_index()['neto']

    # Gross - balance at the end of each month
    summary['Gross'] = original_data.groupby('Month')['updated_balance'].last().reset_index()['updated_balance']

    # Yield Percentage - calculate the percentage yield
    summary['Yield Percentage'] = ((summary['Gross'] - summary['Gross'].shift().fillna(100000)) / summary[
        'Gross'].shift().fillna(100000)) * 100

    # Fill NaN values with 0
    summary = summary.fillna(0)

    # Open the original file in write mode and add a new worksheet named 'Summary'
    with pd.ExcelWriter('data_analyzed.xlsx', engine='openpyxl', mode='a') as writer:
        summary.to_excel(writer, sheet_name='Summary', index=False)

    # Create a line chart
    plt.figure(figsize=(10, 6))
    plt.plot(summary['Month'], summary['Gross'], marker='o', linestyle='-', color='blue')
    plt.title('Monthly Gross Balance')
    plt.xlabel('Month')
    plt.ylabel('Gross Balance')
    plt.grid(True)
    plt.draw()
    plt.pause(0.001)  # Add a short pause to allow the plot to be drawn

    print("Summary and Chart added to the 'data_analyzed.xlsx' file.")


# File: Type_Distribution.py
def analyze_type_distribution():
    # Load the original dataset from the 'Data' worksheet of the XLSX file
    original_data = pd.read_excel('data_analyzed.xlsx', sheet_name='Data')
    original_data['date'] = pd.to_datetime(original_data['date'], format='%d/%m/%Y', errors='coerce')

    # Ensure 'date' is recognized as a datetime column
    original_data['Month'] = original_data['date'].dt.month

    # Create a Type Distribution DataFrame
    type_distribution = pd.DataFrame()

    # Unique types and pl values
    unique_types = original_data['type'].unique()
    unique_pl = original_data['pl'].unique()

    # Create a list to store type/pl combinations and corresponding counts
    type_pl_combinations = []
    type_pl_counts = []

    # Iterate through unique types and pl values
    for t in unique_types:
        for p in unique_pl:
            # Filter data for the current type and pl
            subset = original_data[(original_data['type'] == t) & (original_data['pl'] == p)]

            # Count the occurrences for the current type and pl
            current_count = len(subset)

            # Append the results to the lists
            type_pl_combinations.append((t, p))
            type_pl_counts.append(current_count)

    # Create the Type Distribution DataFrame
    type_distribution['type'], type_distribution['pl'] = zip(*type_pl_combinations)
    type_distribution['Count'] = type_pl_counts

    # Pivot the DataFrame to have types as rows, pl as columns, and corresponding counts in the table
    type_distribution_pivot = type_distribution.pivot_table(index='type', columns='pl', values='Count', fill_value=0)

    # Add a new worksheet named 'Type Distribution' and write the DataFrame
    with pd.ExcelWriter('data_analyzed.xlsx', engine='openpyxl', mode='a') as writer:
        type_distribution_pivot.to_excel(writer, sheet_name='Type Distribution')

    # Create a bar chart using pandas
    plt.figure(figsize=(8, 6))
    type_distribution_pivot.plot(kind='bar', stacked=True, ax=plt.gca())
    plt.title('Type Distribution')
    plt.xlabel('Type')
    plt.ylabel('Count')
    plt.legend(title='PL')
    plt.grid(True)
    plt.draw()
    plt.pause(0.001)  # Add a short pause to allow the plot to be drawn

    print("Type Distribution analysis completed and added to 'data_analyzed.xlsx'")


# File: Hit_By_Symbol.py
def analyze_hit_by_symbol():
    # Load the original dataset from the 'Data' worksheet of the XLSX file
    original_data = pd.read_excel('data_analyzed.xlsx', sheet_name='Data')
    original_data['date'] = pd.to_datetime(original_data['date'], format='%d/%m/%Y', errors='coerce')

    # Ensure 'date' is recognized as a datetime column
    original_data['Month'] = original_data['date'].dt.month

    # Create a Hit By Symbol DataFrame
    hit_by_symbol = pd.DataFrame()

    # Unique symbols and pl values
    unique_symbols = original_data['symbol'].unique()

    # Create lists to store symbol/pl combinations and corresponding counts
    symbol_pl_combinations = []
    symbol_pl_profits = []
    symbol_pl_losses = []
    symbol_pl_total_positions = []

    # Iterate through unique symbols
    for symbol in unique_symbols:
        # Filter data for the current symbol
        subset = original_data[original_data['symbol'] == symbol]

        # Count the occurrences for pl P and pl L for the current symbol
        profits_count = len(subset[subset['pl'] == 'P'])
        losses_count = len(subset[subset['pl'] == 'L'])

        # Calculate total positions, hit percentage, and append the results to the lists
        total_positions = profits_count + losses_count
        hit_percentage = profits_count / total_positions if total_positions > 0 else 0

        symbol_pl_combinations.append(symbol)
        symbol_pl_profits.append(profits_count)
        symbol_pl_losses.append(losses_count)
        symbol_pl_total_positions.append(total_positions)

    # Create the Hit By Symbol DataFrame
    hit_by_symbol['Symbol'] = symbol_pl_combinations
    hit_by_symbol['Profits'] = symbol_pl_profits
    hit_by_symbol['Losses'] = symbol_pl_losses
    hit_by_symbol['Total Positions'] = symbol_pl_total_positions
    hit_by_symbol['Hit Percentage'] = hit_by_symbol['Profits'] / hit_by_symbol['Total Positions']

    # Add a new worksheet named 'Hit By Symbol' and write the DataFrame
    with pd.ExcelWriter('data_analyzed.xlsx', engine='openpyxl', mode='a') as writer:
        hit_by_symbol.to_excel(writer, sheet_name='Hit By Symbol', index=False)

    print("Hit By Symbol analysis completed and added to 'data_analyzed.xlsx'")
    print(hit_by_symbol)


# Main execution
if __name__ == "__main__":
    analyze_data()
    create_summary()
    analyze_type_distribution()
    analyze_hit_by_symbol()

    print("All analyses completed and results added to 'data_analyzed.xlsx'")

    # Show all plots and wait for user to close them
    plt.show()

    print("Program finished. You can close the plot windows to exit.")