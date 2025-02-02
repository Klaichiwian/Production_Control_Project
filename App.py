import pandas as pd
import tkinter as tk
from tkinter import filedialog
from datetime import datetime, timedelta

def browse_file():
    # Open file dialog to choose the input file
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
    return file_path

def adjust_to_shift(start_time):
    # Define working shifts
    shifts = [(datetime.strptime("07:30", "%H:%M").time(), datetime.strptime("16:30", "%H:%M").time()),
              (datetime.strptime("19:30", "%H:%M").time(), datetime.strptime("04:30", "%H:%M").time())]
    
    while True:
        start_time_time = start_time.time()
        for shift_start, shift_end in shifts:
            if shift_start <= shift_end:  # Normal shift (07:30-16:30)
                if shift_start <= start_time_time <= shift_end:
                    return start_time
            else:  # Overnight shift (19:30-04:30)
                if start_time_time >= shift_start or start_time_time <= shift_end:
                    return start_time
        
        # If not in a valid shift, move back one hour and check again
        start_time -= timedelta(hours=1)

def process_bom(input_file, main_part_number, amount, due_datetime, output_file):
    # Load the BOM file
    bom_df = pd.read_excel(input_file)
    
    # Create a dictionary to store the required start times and quantities of each part
    required_parts = {}
    
    # Function to recursively find sub-parts and calculate quantities and start times
    def find_sub_parts(part_number, qty, due_time):
        # Find all rows where the main_partnumber matches
        sub_parts = bom_df[bom_df['main_partnumber'] == part_number]
        
        for _, row in sub_parts.iterrows():
            sub_part = row['sub_partnumber']
            sub_qty = row['Sub Qty'] * qty  # Multiply by the parent quantity
            lead_time = row['Lead Time (sec)']  # Get lead time in seconds
            start_time = due_time - timedelta(seconds=lead_time)
            start_time = adjust_to_shift(start_time)  # Adjust start time to valid shift
            
            # Handle sub-part reuse for multiple main parts
            if sub_part in required_parts:
                # Add the quantity for this part
                required_parts[sub_part]['quantity'] += sub_qty
                # Take the earliest start time for the sub-part, as it might be used in multiple places
                required_parts[sub_part]['start_time'] = min(required_parts[sub_part]['start_time'], start_time)
            else:
                required_parts[sub_part] = {'quantity': sub_qty, 'start_time': start_time}
            
            # Recursive call to find sub-parts of the current sub-part
            find_sub_parts(sub_part, sub_qty, start_time)
    
    # Start processing from the input main part number
    find_sub_parts(main_part_number, amount, due_datetime)
    
    # Convert the result to a DataFrame
    result_df = pd.DataFrame([(k, v['quantity'], v['start_time'].strftime('%d/%m/%Y %H:%M')) for k, v in required_parts.items()], 
                              columns=['Part Number', 'Total Quantity', 'Start Time'])
    
    # Save to Excel
    result_df.to_excel(output_file, index=False)
    print(f"BOM processed and saved to {output_file}")

# Example usage
input_file = browse_file()
if input_file:
    due_datetime = datetime.strptime("10/02/2025 11:30", "%d/%m/%Y %H:%M")
    process_bom(input_file, '208-53-14540', 1, due_datetime, 'Processed_BOM.xlsx')
