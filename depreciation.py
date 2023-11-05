import pandas as pd

def straight_line_depreciation(cost, residual_value, useful_life):
    """Calculate annual depreciation for straight-line method."""
    annual_depreciation = (cost - residual_value) / useful_life
    return annual_depreciation

def generate_straight_line_schedule(cost, residual_value, useful_life):
    """Generate a straight-line depreciation schedule for an asset."""
    annual_depreciation = straight_line_depreciation(cost, residual_value, useful_life)
    schedule = []
    accumulated_depreciation = 0
    
    for year in range(1, useful_life + 1):
        accumulated_depreciation += annual_depreciation
        book_value = cost - accumulated_depreciation
        schedule.append({
            'Year': year,
            'Depreciation Expense': annual_depreciation,
            'Accumulated Depreciation': accumulated_depreciation,
            'Book Value': book_value
        })
    
    return schedule


def declining_balance_depreciation(cost, residual_value, useful_life, factor):
    """Calculate depreciation schedule for declining balance method."""
    schedule = []
    book_value = cost
    for year in range(1, useful_life + 1):
        if book_value < residual_value:
            depreciation_expense = 0
        else:
            depreciation_expense = (factor / useful_life) * book_value
            if depreciation_expense + residual_value > book_value:
                depreciation_expense = book_value - residual_value
        book_value -= depreciation_expense
        schedule.append({
            'Year': year,
            'Depreciation Expense': depreciation_expense,
            'Accumulated Depreciation': cost - book_value,
            'Book Value': book_value
        })
    return schedule

def units_of_production_depreciation(cost, residual_value, total_units, units_used):
    """Calculate depreciation schedule for units of production method."""
    depreciation_per_unit = (cost - residual_value) / total_units
    schedule = []
    accumulated_depreciation = 0
    for year, units in enumerate(units_used, 1):
        depreciation_expense = units * depreciation_per_unit
        accumulated_depreciation += depreciation_expense
        book_value = cost - accumulated_depreciation
        schedule.append({
            'Year': year,
            'Depreciation Expense': depreciation_expense,
            'Accumulated Depreciation': accumulated_depreciation,
            'Book Value': book_value
        })
    return schedule

def generate_depreciation_schedule(method, cost, residual_value, useful_life, factor=None, total_units=None, units_used=None):
    """Generate a depreciation schedule based on selected method."""
    if method == 'straight_line':
        annual_depreciation = straight_line_depreciation(cost, residual_value, useful_life)
        schedule = [{'Year': year, 'Depreciation Expense': annual_depreciation} for year in range(1, useful_life + 1)]
    elif method == 'declining_balance':
        if not factor:
            factor = 2
        schedule = declining_balance_depreciation(cost, residual_value, useful_life, factor)
    elif method == 'units_of_production':
        if total_units and units_used:
            schedule = units_of_production_depreciation(cost, residual_value, total_units, units_used)
        else:
            raise ValueError("Total units and units used must be provided for units-of-production method.")
    else:
        raise ValueError("Invalid depreciation method. Choose 'straight_line', 'declining_balance', or 'units_of_production'.")
    
    if method == 'straight_line':
        accumulated_depreciation = 0
        for year_data in schedule:
            accumulated_depreciation += year_data['Depreciation Expense']
            year_data['Accumulated Depreciation'] = accumulated_depreciation
            year_data['Book Value'] = cost - accumulated_depreciation
    
    return schedule

def write_depreciation_schedules_to_excel(filename):
    """Write depreciation schedules to an Excel file."""
    try:
        # Generate schedules
        schedule_sl = generate_depreciation_schedule('straight_line', 10000, 1000, 5)
        schedule_db = generate_depreciation_schedule('declining_balance', 10000, 1000, 5)
        annual_units_used = [5000, 10000, 15000, 10000, 5000]
        schedule_uop = generate_depreciation_schedule('units_of_production', 10000, 1000, 5, total_units=50000, units_used=annual_units_used)
        
        # Create pandas dataframes
        df_sl = pd.DataFrame(schedule_sl)
        df_db = pd.DataFrame(schedule_db)
        df_uop = pd.DataFrame(schedule_uop)
        
        # Write to Excel
        with pd.ExcelWriter(filename) as writer:
            df_sl.to_excel(writer, sheet_name='Straight_Line')
            df_db.to_excel(writer, sheet_name='Declining_Balance')
            df_uop.to_excel(writer, sheet_name='Units_of_Production')
        
        print(f"Depreciation schedules written to {filename}")
    
    except Exception as e:
        print(f"An error occurred: {e}")

# Call the function to write to Excel
write_depreciation_schedules_to_excel('depreciation_schedules.xlsx')

def sum_of_years_digits_depreciation(cost, residual_value, useful_life):
    """
    Calculates depreciation schedule for the Sum-of-the-Years'-Digits method.
    """
    total_years = sum(range(1, useful_life + 1))
    schedule = []
    accumulated_depreciation = 0
    for year in range(1, useful_life + 1):
        fraction = (useful_life - year + 1) / total_years
        depreciation_expense = (cost - residual_value) * fraction
        accumulated_depreciation += depreciation_expense
        book_value = cost - accumulated_depreciation
        schedule.append({
            'Year': year,
            'Depreciation Expense': depreciation_expense,
            'Accumulated Depreciation': accumulated_depreciation,
            'Book Value': book_value
        })
    return schedule

# Assuming MACRS 5-year property for demonstration purposes (simplified)
def macrs_depreciation(cost, macrs_percentage=[20, 32, 19.2, 11.52, 11.52, 5.76]):
    """
    Calculates depreciation schedule for the MACRS method (5-year property example).
    """
    schedule = []
    accumulated_depreciation = 0
    for year, percentage in enumerate(macrs_percentage, 1):
        depreciation_expense = cost * (percentage / 100)
        accumulated_depreciation += depreciation_expense
        book_value = max(cost - accumulated_depreciation, 0)
        schedule.append({
            'Year': year,
            'Depreciation Expense': depreciation_expense,
            'Accumulated Depreciation': accumulated_depreciation,
            'Book Value': book_value
        })
    return schedule

def partial_year_depreciation(depreciation_schedule, acquisition_month):
    """
    Adjusts depreciation schedule for partial year acquisition.
    """
    # Assuming depreciation starts in the month of acquisition
    # with straight-line proration for the first year.
    monthly_depreciation = depreciation_schedule[0]['Depreciation Expense'] / 12
    months_depreciated = 12 - acquisition_month + 1
    adjusted_first_year_depreciation = monthly_depreciation * months_depreciated

    # Adjust first year depreciation and book value
    depreciation_schedule[0]['Depreciation Expense'] = adjusted_first_year_depreciation
    depreciation_schedule[0]['Book Value'] -= adjusted_first_year_depreciation - depreciation_schedule[0]['Depreciation Expense']

    # Adjust accumulated depreciation
    for i in range(1, len(depreciation_schedule)):
        depreciation_schedule[i]['Accumulated Depreciation'] = (
            depreciation_schedule[i-1]['Accumulated Depreciation'] + depreciation_schedule[i]['Depreciation Expense']
        )
        depreciation_schedule[i]['Book Value'] = (
            depreciation_schedule[i-1]['Book Value'] - depreciation_schedule[i]['Depreciation Expense']
        )

    return depreciation_schedule

def inflation_adjusted_depreciation(depreciation_schedule, inflation_rate):
    """
    Adjusts depreciation schedule for inflation.
    """
    # Adjust book value and depreciation expense for inflation
    for year_data in depreciation_schedule:
        year_data['Depreciation Expense'] *= (1 + inflation_rate) ** (year_data['Year'] - 1)
        year_data['Book Value'] = year_data['Book Value'] * (1 + inflation_rate) ** year_data['Year']
        year_data['Accumulated Depreciation'] = (year_data['Book Value'] - year_data['Depreciation Expense'])

    return depreciation_schedule

def tax_savings_from_depreciation(depreciation_schedule, tax_rate):
    """
    Calculates tax savings from depreciation.
    """
    for year_data in depreciation_schedule:
        year_data['Tax Savings'] = year_data['Depreciation Expense'] * tax_rate

    return depreciation_schedule

def main():
    # Static values for laundromat assets
    assets = {
        'Washing Machines': {'cost': 50000, 'residual_value': 5000, 'useful_life': 7},
        'Furniture and Fixtures': {'cost': 10000, 'residual_value': 0, 'useful_life': 5},
        'Electronics': {'cost': 5000, 'residual_value': 0, 'useful_life': 3},
        'Building Improvements': {'cost': 20000, 'residual_value': 0, 'useful_life': 15},
    }

    # Initialize Excel writer
    filename = 'Laundromat_Depreciation_Schedules.xlsx'
    with pd.ExcelWriter(filename) as writer:
        # Generate and write depreciation schedules for each asset to Excel
        for asset_name, params in assets.items():
            print(f"\nGenerating straight-line depreciation schedule for {asset_name}...")
            schedule = generate_straight_line_schedule(
                params['cost'],
                params['residual_value'],
                params['useful_life']
            )
            
            # Convert to DataFrame
            df_schedule = pd.DataFrame(schedule)
            
            # Write the DataFrame to a specific sheet in the Excel workbook
            df_schedule.to_excel(writer, sheet_name=asset_name, index=False)
            
            # Print confirmation message
            print(f"{asset_name} depreciation schedule written to {filename} in the sheet '{asset_name}'.")

    print(f"\nAll depreciation schedules have been written to {filename}.")

# Ensure the script execution starts here
if __name__ == "__main__":
    main()
