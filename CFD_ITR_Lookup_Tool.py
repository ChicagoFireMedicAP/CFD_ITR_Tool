# In-service training record tool, born out of Illinois State audit of random records
# We moved ITRs to Vector Solutions but have 10+ years of data we may have to access in the future
# legacy data is located on a share point list
# that file is located in private install files or can be downloaded from the source
import pandas as pd
import tkinter as tk
from tkinter import filedialog, simpledialog
import os
import re



# Lets Go!  ITR tracker v3 2/8/24 Anthony.Popelka@cityofchicago.org
#clean and convert duration strings to numbers
def clean_and_convert_duration(duration):
    # Convert the duration to a string
    duration_str = str(duration)
    numeric_part = re.search(r'\d+\.*\d*', duration_str)
    if numeric_part:
        return float(numeric_part.group())
    return 0.0


# prompt user to select a directory for saving the output Excel file
def select_output_directory():
    output_directory = filedialog.askdirectory(title="Select Output Directory")
    return output_directory


# search and destroy
def search_and_write_to_excel():
    file_number = simpledialog.askstring("Enter File Number", "Please enter the File Number:")

    if file_number is not None:
        # Filter the DataFrame based on  file number
        filtered_df = df[
            df[['Officer\'s File #', 'Member 2 File', 'Member 3 File', 'Member 4 File', 'Member 5 File']].eq(
                file_number).any(axis=1)]

        if not filtered_df.empty:
            # store class hours
            class_hours = {}

            # store class hours for each year
            class_hours_by_year = {}

            # Iterate through each row in the df
            for index, row in filtered_df.iterrows():
                year = row['Date'].year

                # Iterate through each drill column and extract drill names and durations
                for i in range(1, 5):
                    drill_column = f"Drill #{i}"
                    duration_column = f"Drill #{i} Duration"
                    drill = row[drill_column]
                    duration = row[duration_column]
                    if not pd.isnull(drill):  # Check if drill exists derp
                        if drill in class_hours:
                            class_hours[drill] += clean_and_convert_duration(duration)
                        else:
                            class_hours[drill] = clean_and_convert_duration(duration)

                        if year in class_hours_by_year:
                            if drill in class_hours_by_year[year]:
                                class_hours_by_year[year][drill] += clean_and_convert_duration(duration)
                            else:
                                class_hours_by_year[year][drill] = clean_and_convert_duration(duration)
                        else:
                            class_hours_by_year[year] = {drill: clean_and_convert_duration(duration)}

            #  dictionary to DataFrame
            class_summary_df = pd.DataFrame(list(class_hours.items()), columns=['Class', 'Total Hours'])

            # Sort the DataFrame 
            class_summary_df = class_summary_df.sort_values(by='Total Hours', ascending=False)

            # output file path
            output_directory = select_output_directory()
            if output_directory:
                output_file_path = os.path.join(output_directory, f"{file_number}_ITR_Records.xlsx")

               
                with pd.ExcelWriter(output_file_path, engine='xlsxwriter') as writer:
                    # class summary to a new sheet
                    class_summary_df.to_excel(writer, sheet_name='Class Summary', index=False)

                    #  class details to the main sheet
                    filtered_df.drop(columns=['Item Type', 'Path'], errors='ignore').to_excel(writer, sheet_name='ITRs', index=False)


                    # separate sheets for each year
                    years = filtered_df['Date'].dt.year.unique()
                    for year in years:
                        year_df = pd.DataFrame(list(class_hours_by_year[year].items()),
                                               columns=['Class', 'Total Hours'])
                        year_df = year_df.sort_values(by='Total Hours', ascending=False)
                        year_sheet_name = str(year)
                        year_df.to_excel(writer, sheet_name=year_sheet_name, index=False)

                print(f"Class details saved to {output_file_path}")
        else:
            print(f"No training records found for File Number: {file_number}")


# GUI window for selecting the Excel file
root = tk.Tk()
root.withdraw()

# select the Excel file
selected_file = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])

if selected_file:
    df = pd.read_excel(selected_file, sheet_name='ITRs')  # Assuming the sheet name is 'ITRs'

    # GUI window
    root = tk.Tk()
    root.title("Training Record Lookup")

    # Center the window
    window_width = 400
    window_height = 150
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    x_position = (screen_width - window_width) // 2
    y_position = (screen_height - window_height) // 2
    root.geometry(f"{window_width}x{window_height}+{x_position}+{y_position}")

    #  button to trigger the search and write to Excel
    search_button = tk.Button(root, text="Search and Write to Excel", command=search_and_write_to_excel)
    search_button.pack(pady=20)

    # Run the GUI loop
    root.mainloop()
else:
    print("No Excel file selected. Exiting program.")

# End 
