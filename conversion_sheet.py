"""
Conversion Tracker Application

This application helps track customer conversion rates by calculating the number
of customers from morning and closing counts and comparing with transaction numbers.
Data is saved to an Excel spreadsheet for record keeping.
"""

import os
import tkinter as tk
from tkinter import messagebox
from datetime import datetime
import openpyxl


# CONSTANTS
PADDING = 50
ENTRY_WIDTH = 15
TEXT_HEIGHT = 5
TEXT_WIDTH = 20

# FONT CONSTANTS
PERSONAL_NOTE_COLOR = 'green'
PERSONAL_NOTE_FONT = ('Helvetica', 13, 'bold')
STANDARD_FONT = ('Arial', 11, 'bold')
COMMENT_FONT = ('Arial', 9)


class ConversionTracker:
    def __init__(self, root):
        self.root = root
        self.setup_window()
        self.create_widgets()
        
    def setup_window(self):
        """Configure the main window"""
        self.root.title("Conversion Tracker")
        self.root.config(padx=PADDING, pady=PADDING)
        
    def create_widgets(self):
        """Create all UI widgets"""
        self.create_labels()
        self.create_entries()
        self.create_buttons()
        
    def create_labels(self):
        """Create all label widgets"""
        labels = [
            "Day:", "Date:", "Morning Count:", "Closing Count:", "# Of Cust.:",
            "# Of Trans.:", "Conv. Rate:", "END of DAY (sales w/o tax):", "Comments:"
        ]
        
        for idx, text in enumerate(labels):
            label = tk.Label(self.root, text=text, font=STANDARD_FONT)
            label.grid(column=0, row=idx, pady=5, sticky=tk.W)
            
    def create_entries(self):
        """Create all entry widgets"""
        # Day dropdown
        self.day_var = tk.StringVar(value="Mon")
        days = ["Mon", "Tues", "Wed", "Thurs", "Fri", "Sat", "Sun"]
        self.day_menu = tk.OptionMenu(self.root, self.day_var, *days)
        self.day_menu.grid(column=1, row=0)
        
        # Date entry with default current date
        self.date_entry = tk.Entry(self.root, width=ENTRY_WIDTH)
        self.date_entry.grid(column=1, row=1)
        self.date_entry.insert(0, datetime.now().strftime("%m-%d-%Y"))
        
        # Count entries
        self.morning_count_entry = tk.Entry(self.root, width=ENTRY_WIDTH)
        self.morning_count_entry.grid(column=1, row=2)
        
        self.closing_count_entry = tk.Entry(self.root, width=ENTRY_WIDTH)
        self.closing_count_entry.grid(column=1, row=3)
        
        # Results entries (disabled for manual entry)
        self.customers_entry = tk.Entry(self.root, width=ENTRY_WIDTH, state=tk.DISABLED)
        self.customers_entry.grid(column=1, row=4)
        
        self.transaction_number_entry = tk.Entry(self.root, width=ENTRY_WIDTH)
        self.transaction_number_entry.grid(column=1, row=5)
        
        self.conversion_rate_entry = tk.Entry(self.root, width=ENTRY_WIDTH, state=tk.DISABLED)
        self.conversion_rate_entry.grid(column=1, row=6)
        
        # End of day entry
        self.end_of_day_entry = tk.Entry(self.root, width=ENTRY_WIDTH)
        self.end_of_day_entry.grid(column=1, row=7)
        
        # Comments text area
        self.comments_entry = tk.Text(self.root, height=TEXT_HEIGHT, width=TEXT_WIDTH, font=COMMENT_FONT)
        self.comments_entry.grid(column=1, row=8, columnspan=2)
        
    def create_buttons(self):
        """Create all button widgets"""
        # Calculate button
        self.calculate_button = tk.Button(self.root, text="Calculate", command=self.calculate_and_fill)
        self.calculate_button.grid(column=2, row=9, pady=10)
        
        # Confirm button (hidden initially)
        self.confirm_button = tk.Button(self.root, text="Confirm", command=self.confirm_and_save)
        self.confirm_button.grid(column=2, row=10, pady=10)
        self.confirm_button.grid_remove()
        
    def calculate_number_of_customers(self, closing_count, morning_count):
        """Calculate number of customers by subtracting morning count from closing count"""
        return closing_count - morning_count
        
    def calculate_conversion_rate(self, transaction_number, number_of_customers):
        """Calculate conversion rate as percentage of transactions to customers"""
        if number_of_customers > 0:
            return f'{transaction_number / number_of_customers * 100:.2f}%'
        return "0.00%"
        
    def calculate_and_fill(self):
        """Calculate values based on user input and fill in results fields"""
        try:
            morning_count = int(self.morning_count_entry.get())
            closing_count = int(self.closing_count_entry.get())
            transaction_number = int(self.transaction_number_entry.get())
            
            number_of_customers = self.calculate_number_of_customers(closing_count, morning_count)
            conversion_rate = self.calculate_conversion_rate(transaction_number, number_of_customers)
            
            # Update customer entry
            self.customers_entry.config(state=tk.NORMAL)
            self.customers_entry.delete(0, tk.END)
            self.customers_entry.insert(0, number_of_customers)
            self.customers_entry.config(state=tk.DISABLED)
            
            # Update conversion rate entry
            self.conversion_rate_entry.config(state=tk.NORMAL)
            self.conversion_rate_entry.delete(0, tk.END)
            self.conversion_rate_entry.insert(0, conversion_rate)
            self.conversion_rate_entry.config(state=tk.DISABLED)
            
            # Show confirm button
            self.confirm_button.grid()
        except ValueError:
            messagebox.showerror("Input Error", "Please enter valid numbers for counts and transactions.")
    
    def save_to_excel(self, data):
        """Save data to Excel spreadsheet"""
        file_name = 'conversion_data.xlsx'
        try:
            workbook = openpyxl.load_workbook(file_name)
            sheet = workbook.active
        except FileNotFoundError:
            workbook = openpyxl.Workbook()
            sheet = workbook.active
            sheet.append(['Date', 'Day', 'Morning Count', 'Closing Count', 'Transaction Number',
                         'Number of Customers', 'Conversion Rate', 'End of Day', 'Comments'])
        
        sheet.append([
            data['date'], data['day'], data['morning_count'], data['closing_count'], 
            data['transaction_number'], data['number_of_customers'], data['conversion_rate'],
            data['end_of_day'], data['comments']
        ])
        
        workbook.save(file_name)
        
        # Open the Excel file to show the user
        try:
            os.startfile(file_name)  # Windows
        except AttributeError:
            os.system(f'open "{file_name}"')  # macOS
        except Exception as e:
            messagebox.showerror("Error", f"Could not open file: {e}")
    
    def confirm_and_save(self):
        """Confirm the entries and save data to Excel"""
        try:
            # Collect all data
            data = {
                'date': self.date_entry.get(),
                'day': self.day_var.get(),
                'morning_count': int(self.morning_count_entry.get()),
                'closing_count': int(self.closing_count_entry.get()),
                'transaction_number': int(self.transaction_number_entry.get()),
                'number_of_customers': int(self.customers_entry.get()),
                'conversion_rate': self.conversion_rate_entry.get(),
                'end_of_day': int(self.end_of_day_entry.get()),
                'comments': self.comments_entry.get("1.0", tk.END).strip()
            }
            
            # Save data to Excel
            self.save_to_excel(data)
            messagebox.showinfo("Success", "Data saved successfully!")
            
            # Reset fields
            self.reset_fields()
        except ValueError:
            messagebox.showerror("Error", "Please ensure all fields are filled correctly.")
    
    def reset_fields(self):
        """Reset all entry fields to their default state"""
        # Reset date to current date
        self.date_entry.delete(0, tk.END)
        self.date_entry.insert(0, datetime.now().strftime("%m-%d-%Y"))
        
        # Clear count entries
        self.morning_count_entry.delete(0, tk.END)
        self.closing_count_entry.delete(0, tk.END)
        self.transaction_number_entry.delete(0, tk.END)
        
        # Clear and disable calculated entries
        self.customers_entry.config(state=tk.NORMAL)
        self.customers_entry.delete(0, tk.END)
        self.customers_entry.config(state=tk.DISABLED)
        
        self.conversion_rate_entry.config(state=tk.NORMAL)
        self.conversion_rate_entry.delete(0, tk.END)
        self.conversion_rate_entry.config(state=tk.DISABLED)
        
        # Clear other entries
        self.end_of_day_entry.delete(0, tk.END)
        self.comments_entry.delete("1.0", tk.END)
        
        # Reset dropdown
        self.day_var.set("Mon")
        
        # Hide confirm button
        self.confirm_button.grid_remove()


# Main application entry point
if __name__ == "__main__":
    root = tk.Tk()
    app = ConversionTracker(root)
    root.mainloop()
