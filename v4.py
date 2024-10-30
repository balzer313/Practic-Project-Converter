from datetime import datetime
import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import docx
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.shared import Pt

import json
import pandas as pd
from PIL import Image, ImageTk

def convert_time(timestamp):
    return datetime.utcfromtimestamp(timestamp / 1000).strftime('%b %d, %Y')
def discount_breakdown(adjustments):
    for adjustment in adjustments:
        if adjustment.get('promoCode') is not None and adjustment.get('type') == 'PROMO_CODE':
            percent = adjustment.get('percent', 0)
            promo_code = adjustment.get('promoCode')
            return f"{percent}% PROMO_CODE: {promo_code}"
    return "No Discounts"
def select_json_file():
    global selected_json_file
    selected_json_file = filedialog.askopenfilename(title="Select Json File",
                                                    filetypes=(("Json Files", "*.json"), ("All Files", "*.*")))
    if not selected_json_file:
        return  # User canceled file selection
    display_selected_file()
def display_selected_file():#screen 2
    global frame2  # Declare new_frame as a global variable
    global generate_word_file
    frame.grid_forget()  # Remove the existing frame
    frame2 = ttk.Frame(root, padding=20)
    frame2.grid(row=0, column=0)

    new_title_label = ttk.Label(frame2, text="Selected Excel File", font=("Arial", 16, "bold"))
    new_title_label.grid(row=0, column=0, pady=(0, 20), columnspan=2)

    excel_file_label = ttk.Label(frame2, text=selected_json_file, font=("Arial", 10))
    excel_file_label.grid(row=1, column=0, pady=(0, 10), padx=10, sticky='w')

    change_excel_button = ttk.Button(frame2, text="Change Json File", command=change_selected_file)
    change_excel_button.grid(row=2, column=0, pady=(10, 10), padx=10, sticky='w')

    generate_word_checkbox = ttk.Checkbutton(frame2, text="Generate Word File", variable=generate_word_file)
    generate_word_checkbox.grid(row=3, column=0, pady=(10, 10), padx=10, sticky='w')

    convert_button = ttk.Button(frame2, text="Convert", command=convert)
    convert_button.grid(row=4, column=0, pady=(10, 10), padx=10, sticky='w')

    # Center all widgets both vertically and horizontally
    frame2.grid_rowconfigure(1, weight=1)
    frame2.grid_columnconfigure(0, weight=1)

def change_selected_file():
    global selected_json_file
    global frame2  # Make new_frame accessible
    selected_json_file = None
    frame2.grid_forget()  # Remove the existing frame
    frame.grid()  # Re-display the original frame to select a new file2

def convert():

    global selected_json_file
    output_folder = filedialog.askdirectory(title="Select Folder to Save File")
    if output_folder:
        try:
            # Load the JSON file
            with open(selected_json_file, 'r') as f:
                data = json.load(f)

            # Extract relevant data from JSON
            quote_details = data.get('QuoteDetails', {})
            line_items = quote_details.get('upcomingBills', {}).get('lines', [])
            for_users = data.get("QuoteDetails", {}).get("lineItems", [])

            # Create a list to store the table data
            table_data = []


            annu_month = "MONTHLY"  # set the deafult to monthly type
            # Extract details for each line item
            for item in line_items:
                #formatted_number = "{:,.2f}".format(number) - format numbers with ,
                line_item = {}
                line_item['Product'] = f"{item['description']}, up to {item['quantity']} users"  # Use description for Product
                if "ANNUAL" in item['description']: annu_month = "ANNUAL"  # checks if annual or monthly
                line_item['List Price'] = "{:,.2f}".format(item['subTotal']/100)  # Use subTotal for List Price
                line_item['Discount'] = "{:,.2f}".format((item['subTotal'] - item['total'] - (item['tax'] if 'tax' in item else 0))/100)
                line_item['Amount excl. tax'] = "{:,.2f}".format((item['total'] + item['tax'] if 'tax' in item else 0)/100)
                line_item['Tax'] = "{:,.2f}".format((item['tax'] if 'tax' in item else 0)/100)  # Use tax if available, otherwise 0
                line_item['Amount (USD)'] = "{:,.2f}".format(item['total']/100)  # Use total for Amount
                line_item['Billing period'] = f"{convert_time(item['period']['startsAt'])} - {convert_time(item['period']['endsAt'])}"

                # line_item['Users'] = item['quoteLineId']
                for item_up in for_users:
                    if item_up.get("lineItemId") == item['quoteLineId']:
                        try:
                            line_item['Users'] = item_up.get("chargeQuantities", [])[0].get("quantity")
                        except Exception as e:
                            line_item['Users'] = None
                    else: line_item['Users'] = None

                line_item['Discounts breakdown'] = discount_breakdown(item.get('adjustments', []))

                not_to_check = ['Product', 'Users', 'Billing period', 'Discounts breakdown']
                all_zero = all(
                    float(value.replace(',', '')) == 0
                    for key, value in line_item.items() if (key not in not_to_check)
                )
                if all_zero:  # checks if all the line is 0
                    continue

                table_data.append(line_item)

            # Convert the table data to a DataFrame
            df = pd.DataFrame(table_data)
            df.columns = ['Product', 'List Price (USD)', 'Discount', 'Amount excl. tax (USD)', 'Tax', 'Amount', 'Billing period', 'Users', 'Discounts breakdown']

            company_name = data['InvoiceGroup']['shipToParty']['name']
            quote_number = data['QuoteDetails']['number']

            # make a word file if checked
            word_file()

            # save the excel
            excel_name = f'{annu_month} Excel - {company_name}.xlsx'  # excel file name
            output_excel = os.path.join(output_folder, excel_name)
            df.to_excel(output_excel, index=False)

            current_time = datetime.now().strftime('%Y-%m-%d, %H:%M:%S')  # Get the current date and time
            messagebox.showinfo("Success", f"Excel file created at {output_excel}\nCompany: {company_name}\nQuote number: {quote_number}\nUpdated: {current_time}")
            root.destroy()

        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {str(e)}")


def word_file(): ## creating(if wonted) a word file
    global generate_word_file
    def change_word(docs, old_word, new_word, bold, size, page_break):
        for paragraph in docs.paragraphs:
            if f'%%{old_word}%%' in paragraph.text:
                # print('ok')
                paragraph.text = paragraph.text.replace(f'%%{old_word}%%',
                                                        new_word)  # Replace old_word with the new_word
                for run in paragraph.runs:
                    if new_word in run.text:
                        run.font.size = Pt(size)  # Set font size to 18pt
                        run.font.bold = bold  # Set text to bold
                        run.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER  # Center align the text
                if page_break:
                    paragraph.clear()
                    run2 = paragraph.add_run()
                    run2.add_break(docx.enum.text.WD_BREAK.PAGE)
    if generate_word_file:
        return ## handle the word file

money_type = "DOLLAR"

def toggle_currency():
    global money_type
    if money_type == "DOLLAR": money_type = "DIRHAM"; toggle_currency_button.configure(image=off_icon)
    else: money_type = "DOLLAR"; toggle_currency_button.configure(image=on_icon)

# Set up the Tkinter window
root = tk.Tk()
root.title("JSON to Excel and Docx Converter")

# Get the screen width and height
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()
window_width = 600
window_height = 300
x = (screen_width - window_width) // 2
y = (screen_height - window_height) // 2
root.geometry(f"{window_width}x{window_height}+{x}+{y}")

# Load images for on and off states
on_image = Image.open("images/usa_flag.png").resize((45, 30))
on_icon = ImageTk.PhotoImage(on_image)
off_image = Image.open("images/Flag-United-Arab-Emirates.png").resize((45, 30))
off_icon = ImageTk.PhotoImage(off_image)

currency_var = tk.BooleanVar(value=True)  # True for DOLLAR, False for DIRHAM
generate_word_file = tk.BooleanVar(value=False)

# Create a frame with padding
frame = ttk.Frame(root, padding=20)
frame.grid(row=0, column=0, sticky="nsew")

# Create a label for the title
title_label = ttk.Label(frame, text="JSON to Excel and Docx Converter", font=("Arial", 16, "bold"))
title_label.grid(row=0, column=0, pady=(0, 20), columnspan=2)

# Create a button to select the JSON file
select_excel_button = ttk.Button(frame, text="Select JSON File", command=select_json_file)
select_excel_button.grid(row=1, column=0, pady=(0, 10), padx=10, sticky='w')

# Label to display the selected file name
json_file_label = ttk.Label(frame, text="", font=("Arial", 10))
json_file_label.grid(row=2, column=0, pady=(10, 10), padx=10, sticky='w')

# Currency toggle button
toggle_currency_button = ttk.Checkbutton(
    frame, image=on_icon, variable=currency_var, command=toggle_currency,
    style="TButton", onvalue=True, offvalue=False
)
toggle_currency_button.grid(row=3, column=0, pady=(0, 10), padx=10, sticky='w')

# Configure window grid to make it responsive
root.grid_rowconfigure(0, weight=1)
root.grid_columnconfigure(0, weight=1)

root.mainloop()
