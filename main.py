import openpyxl
import datetime
import sys  

# Specify the path to your Excel file
excel_file_path = './Assignment_Timecard.xlsx'

# Load the workbook
workbook = openpyxl.load_workbook(excel_file_path)

# Get the first sheet from the workbook
sheet = workbook.active

# Iterate through each row in the sheet
filtered_rows = []
for row in sheet.iter_rows(min_row=2, values_only=True):
    # Store columns 0, 2, 3, and 4 in the 2D list
    if all(cell is not None and cell != '' for cell in [row[0], row[2], row[3], row[4], row[7]]):
        # Store the row if none of the specified columns is empty
        filtered_rows.append([row[0], row[2], row[3], row[4], row[7]])

# Close the workbook
workbook.close()

# Initialize variables for analysis
updated = []


def analyze():
    # Initialize variables to track consecutive days, time between shifts, and single shift duration
    prev = filtered_rows[0]
    day = 1
    seven_cons_days = False
    one_to_ten = False
    greater_than_fourteen = False

    # Iterate through each row in filtered_rows
    for row in filtered_rows[1:]:
        # Check if the employee ID changes, indicating a new employee
        if row[0] != prev[0]:
            # Store analysis results for the previous employee
            updated.append([prev[0], prev[4], seven_cons_days, one_to_ten, greater_than_fourteen])

            # Reset variables for the new employee
            day = 1
            one_to_ten = False
            greater_than_fourteen = False
            seven_cons_days = False
        else:
            # Check if the current row is a different day for the same employee
            if prev[1].date() != row[1].date():
                # Calculate time difference between shifts in hours
                d_time = (row[1] - prev[1]).total_seconds() / 3600.0

                # If the time difference is more than 24 hours, reset the consecutive day counter
                if d_time > 24.0:
                    day = 0
                day += 1

        # Check if the employee has worked for 7 consecutive days
        if day >= 7:
            seven_cons_days = True

        # Extract hours and minutes from the current row's time duration
        current_time_str = row[3]
        hours, minutes = map(int, current_time_str.split(':'))

        # Convert the time duration to timedelta
        current_time = datetime.timedelta(hours=hours, minutes=minutes)
        current_hours = current_time.total_seconds() / 3600.0

        # Check if the employee has worked for more than 14 hours in a single shift
        if current_hours > 14:
            greater_than_fourteen = True

        # Calculate time difference between the end time of the current row and the start time of the previous row
        end_time_row1 = row[1]
        start_time_row2 = prev[2]
        time_difference = end_time_row1 - start_time_row2

        # Extract hours from the time difference
        hours = time_difference.total_seconds() / 3600.0

        # Check if the employee has less than 10 hours of time between shifts but greater than 1 hour
        if 1 < hours < 10:
            one_to_ten = True

        # Update the previous row for the next iteration
        prev = row


# Run the analysis
analyze()


# Function to print the analysis report
def analysis_report():
    print("Id and Name for who has worked for 7 consecutive days.")
    for row in updated:
        if row[2]:
            print(row[0], "   ", row[1])

    print("\nId and Name for who have less than 10 hours of time between shifts but greater than 1 hour")
    for row in updated:
        if row[3]:
            print(row[0], "   ", row[1])

    print("\nId and Name for Who has worked for more than 14 hours in a single shift")
    for row in updated:
        if row[4]:
            print(row[0], "   ", row[1])


# Print the analysis report
analysis_report()

# Redirect console output to a file
with open('output.txt', 'w') as f:
    # Save the original standard output
    original_stdout = sys.stdout
    # Redirect standard output to the file
    sys.stdout = f

    
    # Print the analysis report
    analysis_report()

    # Restore the original standard output
    sys.stdout = original_stdout

print("Output has been saved to 'output.txt'")
