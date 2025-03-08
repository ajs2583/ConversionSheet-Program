# tkinter UI library
from tkinter import *
# Message box UI
from tkinter import messagebox
# date time for current date time
from datetime import datetime
# Library for excel spreadsheets
import openpyxl
# For opening the Excel file
import os


# FONT CONSTANTS
PERSONAL_NOTE_COLOR = 'green'
PERSONAL_NOTE_FONT = ('Helvetica', 13, 'bold')
STANDARD_FONT = ('Arial', 11, 'bold')


#! Backend Logic ###############################################################
'''
This function calculates the number of customers by subtracting the morning count from the closing count.
args -- closing count and morning count (int)
returns: difference between closing and morning count
'''

def calculate_number_of_customers(closing_count: int, morning_count: int) -> int:
    return closing_count - morning_count


def conversion_rate_calculation(transaction_number: int, number_of_customers: int) -> str:
    return f'{transaction_number / number_of_customers * 100:.2f}%' if number_of_customers > 0 else "0.00%"


def save_to_excel(date, day, morning_count, closing_count, transaction_number, number_of_customers, conversion_rate, end_of_day, comments):
    file_name = 'conversion_data.xlsx'
    try:
        workbook = openpyxl.load_workbook(file_name)
    except FileNotFoundError:
        workbook = openpyxl.Workbook()
        sheet.append(['Date', 'Day', 'Morning Count', 'Closing Count', 'Transaction Number',
                      'Number of Customers', 'Conversion Rate', 'End of Day', 'Comments'])

    sheet = workbook.active
    sheet.append([date, day, morning_count, closing_count, transaction_number,
                  number_of_customers, conversion_rate, end_of_day, comments])
    workbook.save(file_name)

    # Open the Excel file to show the user
    try:
        os.startfile(file_name)  # Windows
    except AttributeError:
        os.system(f'open "{file_name}"')
    except Exception as e:
        messagebox.showerror("Error", f"Could not open file: {e}")


def calculate_and_fill():
    try:
        morning_count = int(morning_count_entry.get())
        closing_count = int(closing_count_entry.get())
        transaction_number = int(transaction_number_entry.get())

        number_of_customers = calculate_number_of_customers(closing_count, morning_count)
        conversion_rate = conversion_rate_calculation(transaction_number, number_of_customers)

        no_of_customers_entry.config(state=NORMAL)
        no_of_customers_entry.delete(0, END)
        no_of_customers_entry.insert(0, number_of_customers)
        no_of_customers_entry.config(state=DISABLED)

        conversion_rate_entry.config(state=NORMAL)
        conversion_rate_entry.delete(0, END)
        conversion_rate_entry.insert(0, conversion_rate)
        conversion_rate_entry.config(state=DISABLED)

        confirm_button.grid()

    except ValueError:
        messagebox.showerror("Input Error", "Please enter valid numbers for counts and transactions.")


def confirm_and_save():
    try:
        date = date_entry.get()  # Retrieve the entered date
        day = clicked.get()
        morning_count = int(morning_count_entry.get())
        closing_count = int(closing_count_entry.get())
        transaction_number = int(transaction_number_entry.get())
        number_of_customers = int(no_of_customers_entry.get())
        conversion_rate = conversion_rate_entry.get()
        end_of_day = int(end_of_day_entry.get())  # Retrieve EOD value
        comments = comments_entry.get("1.0", END).strip()  # Get comments and strip extra spaces/newlines

        save_to_excel(date, day, morning_count, closing_count, transaction_number, number_of_customers, conversion_rate, end_of_day, comments)
        messagebox.showinfo("Success", "Data saved successfully!")

        # Clear all fields after confirmation
        date_entry.delete(0, END)
        date_entry.insert(0, datetime.now().strftime("%m-%d-%Y"))  # Reset to current date

        morning_count_entry.delete(0, END)
        closing_count_entry.delete(0, END)
        transaction_number_entry.delete(0, END)

        no_of_customers_entry.config(state=NORMAL)
        no_of_customers_entry.delete(0, END)
        no_of_customers_entry.config(state=DISABLED)

        conversion_rate_entry.config(state=NORMAL)
        conversion_rate_entry.delete(0, END)
        conversion_rate_entry.config(state=DISABLED)

        end_of_day_entry.delete(0, END)
        comments_entry.delete("1.0", END)  # Clear the Text widget

        # Reset dropdown to default value
        clicked.set("Mon")  

    except ValueError:
        messagebox.showerror("Error", "Please ensure all fields are filled correctly.")



# UI Setup
window = Tk()
window.title("Conversion Tracker")
window.config(padx=50, pady=50)

# Labels
labels = [
    "Day:", "Date:", "Morning Count:", "Closing Count:", "# Of Cust.:",
    "# Of Trans.:", "Conv. Rate:", "END of DAY (sales w/o tax):", "Comments:"
]


for idx, text in enumerate(labels):
    label = Label(window, text=text, font=STANDARD_FONT)
    label.grid(column=0, row=idx, pady=5)

# Entries #####################################################################

# Date Entry with Default Current Date
date_entry = Entry(window, width=15)
date_entry.grid(column=1, row=1)
date_entry.insert(0, datetime.now().strftime("%m-%d-%Y"))  # Default to today's date

morning_count_entry = Entry(window, width=15)
morning_count_entry.grid(column=1, row=2)

closing_count_entry = Entry(window, width=15)
closing_count_entry.grid(column=1, row=3)

no_of_customers_entry = Entry(window, width=15, state=DISABLED)
no_of_customers_entry.grid(column=1, row=4)

transaction_number_entry = Entry(window, width=15)
transaction_number_entry.grid(column=1, row=5)

conversion_rate_entry = Entry(window, width=15, state=DISABLED)
conversion_rate_entry.grid(column=1, row=6)

# End of Day (EOD) Entry
end_of_day_entry = Entry(window, width=15)
end_of_day_entry.grid(column=1, row=7)

# Comments Section
comments_entry = Text(window, height=5, width=20, font=('Arial', 9))
comments_entry.grid(column=1, row=8, columnspan=2)

# Dropdown Menu for Day Selection
clicked = StringVar()
clicked.set("Mon")
day_menu = OptionMenu(window, clicked, "Mon", "Tues", "Wed", "Thurs", "Fri", "Sat", "Sun")
day_menu.grid(column=1, row=0)

# Buttons
calculate_button = Button(window, text="Calculate", command=calculate_and_fill)
calculate_button.grid(column=2, row=9, pady=10)

confirm_button = Button(window, text="Confirm", command=confirm_and_save)
confirm_button.grid(column=2, row=10, pady=10)
confirm_button.grid_remove()

# Main loop
window.mainloop()
