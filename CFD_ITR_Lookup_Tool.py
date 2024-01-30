# In-service traiing record tool, born out of Illinois State audit of random records
# We moved ITRs to Vector Solutions but have 10+ years of data we may have to access in the future
import pandas as pd
import tkinter as tk
from tkinter import filedialog, simpledialog, ttk
import os
import re


# Function to clean and convert duration strings to numbers
def clean_and_convert_duration(duration):
    duration_str = str(duration)
    numeric_part = re.search(r'\d+\.*\d*', duration_str)
    if numeric_part:
        return float(numeric_part.group())
    return 0.0


# Function to prompt user to select a directory for saving the output Excel file
def select_output_directory():
    output_directory = filedialog.askdirectory(title="Select Output Directory")
    return output_directory


# Function to search and write class details to an Excel file
def search_and_write_to_excel():
    file_number = simpledialog.askstring("Enter File Number", "Please enter the File Number:")

    if file_number is not None:
        filtered_df = df[
            df[['Officer\'s File #', 'Member 2 File', 'Member 3 File', 'Member 4 File', 'Member 5 File']].eq(
                file_number).any(1)]

        if not filtered_df.empty:
            class_hours = {}

            for index, row in filtered_df.iterrows():
                for i in range(1, 5):
                    drill_column = f"Drill #{i}"
                    duration_column = f"Drill #{i} Duration"
                    drill = row[drill_column]
                    duration = row[duration_column]
                    if not pd.isnull(drill):
                        if drill in class_hours:
                            class_hours[drill] += clean_and_convert_duration(duration)
                        else:
                            class_hours[drill] = clean_and_convert_duration(duration)

            class_summary_df = pd.DataFrame(list(class_hours.items()), columns=['Class', 'Total Hours'])
            class_summary_df = class_summary_df.sort_values(by='Total Hours', ascending=False)

            output_directory = select_output_directory()
            if output_directory:
                output_file_path = os.path.join(output_directory, f"{file_number}_ITR_Records.xlsx")

                with pd.ExcelWriter(output_file_path, engine='xlsxwriter') as writer:
                    class_summary_df.to_excel(writer, sheet_name='Class Summary', index=False)
                    filtered_df.drop(columns=['Item Type', 'Path']).to_excel(writer, sheet_name='ITRs', index=False)

                print(f"Class details saved to {output_file_path}")
        else:
            print(f"No training records found for File Number: {file_number}")


# Function to display instructions in a pop-up window
def display_instructions():
    instructions_window = tk.Toplevel(root)
    instructions_window.title("Instructions")

    instructions_text = """Welcome to the Training Record Lookup Tool!

This tool allows you to search and extract information from an Excel file containing training records.

Here's how to use it:

1. Click the "Select Excel File" button and choose the Excel file with your training records.

2. Enter the File Number in the provided dialog box. This will be used to filter the records.

3. Choose the output directory where the generated Excel file will be saved.

4. Click "Search and Write to Excel" to perform the search and extraction.

Once you've read and understood these instructions, click the "I understand" button to continue.
"""

    instructions_label = tk.Label(instructions_window, text=instructions_text, justify="left")
    instructions_label.pack(padx=20, pady=20)

    def close_instructions():
        instructions_window.destroy()

    # Add an "I understand" button to acknowledge the instructions
    understand_button = tk.Button(instructions_window, text="I understand", command=close_instructions)
    understand_button.pack(pady=10)

    instructions_window.focus_force()
    instructions_window.grab_set()
    instructions_window.wait_window()


# Function to select the Excel file with ITR data
def select_excel_file():
    global df  # Make the DataFrame accessible globally
    selected_file = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])

    if selected_file:
        # Read the data from the selected Excel file into a Pandas DataFrame
        df = pd.read_excel(selected_file, sheet_name='ITRs')  # Assuming the sheet name is 'ITRs'


# Create a tkinter GUI window
root = tk.Tk()
root.title("Training Record Lookup")

# Center the window on the screen
window_width = 400
window_height = 150
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()
x_position = (screen_width - window_width) // 2
y_position = (screen_height - window_height) // 2
root.geometry(f"{window_width}x{window_height}+{x_position}+{y_position}")

# Create a button to display instructions
instructions_button = tk.Button(root, text="Show Instructions", command=display_instructions)
instructions_button.pack(pady=20)

# Create a button to select the Excel file with ITR data
select_excel_button = tk.Button(root, text="1st Step: Select ITR Excel File", command=select_excel_file)
select_excel_button.pack(pady=10)

# Create a button to trigger the search and write to Excel
search_button = tk.Button(root, text="2nd Step: Search File Number and Write the Report", command=search_and_write_to_excel)
search_button.pack(pady=10)

# Initialize a DataFrame
df = None

# Run the GUI loop
root.mainloop()
