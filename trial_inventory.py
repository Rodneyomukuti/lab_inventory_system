import pandas as pd
from tkinter import *
from tkinter import messagebox
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from datetime import datetime
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from gspread_dataframe import set_with_dataframe, get_as_dataframe

# Function to authenticate and connect to the Google Sheet
def connect_to_google_sheet(sheet_name):
    scope = ["https://spreadsheets.google.com/feeds", 'https://www.googleapis.com/auth/drive']
    creds = ServiceAccountCredentials.from_json_keyfile_name('inventory-438319-ae30963a455c.json', scope)
    client = gspread.authorize(creds)
    sheet = client.open(sheet_name).sheet1
    return sheet

# Load the Google Sheet data into a pandas DataFrame
def load_inventory():
    sheet = connect_to_google_sheet('trial')  # Replace this with the original sheet name
    data = sheet.get_all_records()
    df = pd.DataFrame(data)
    df.columns = df.columns.str.strip()  # Removes any extra spaces
    return df

# Save the updated DataFrame to the Google Sheet
def save_inventory(df):
    sheet = connect_to_google_sheet('trial')  # Replace this with the original sheet name
    sheet.clear()  # Clear the sheet before updating
    set_with_dataframe(sheet, df)  # Write the df to the sheet

# Update inventory when a reagent/consumable is stored
def store_inventory():
    date = entry_date.get()
    item_code = entry_item_code.get()
    product_name = entry_product_name.get()
    manufacturer = entry_manufacturer.get()
    received = entry_received.get()
    b_fwd = entry_b_fwd.get()
    qty = int(entry_qty.get())
    unit = entry_unit.get()
    balance_qty = int(entry_balance_qty.get())
    location = entry_location.get()
    reorder_level = int(entry_reorder_level.get())
    comments = entry_comments.get()

    df = load_inventory()

    # Check if product exists in inventory
    if product_name in df['Product name'].values:
        # Update balance and other details for existing product
        df.loc[df['Product name'] == product_name, 'Balance (Qty)'] = balance_qty + qty
    else:
        # Create a new row as a DataFrame
        new_row = pd.DataFrame({
            'Date': [date], 'Item code': [item_code], 'Product name': [product_name], 'Manufacturer': [manufacturer],
            'Received?': [received], 'Issued?': [''], 'B/fwd?': [b_fwd], 'Qty': [qty], 'Unit': [unit],
            'Balance (Qty)': [balance_qty], 'Location (in store)': [location], 'Reorder level': [reorder_level], 
            'Comments': [comments]
        })

        # Use pd.concat() to add the new row to the DataFrame (pd.append doesn't work)
        df = pd.concat([df, new_row], ignore_index=True)

    # Save the updated inventory
    save_inventory(df)
    messagebox.showinfo("Success", "Inventory updated successfully!")

# Handle picking an item (taking from inventory)
def take_inventory():
    # Get input from the user
    date = entry_date.get()  # Get the date
    product_name = entry_product_name.get()  # Get the product name
    item_code = entry_item_code.get()  # Get the item code
    manufacturer = entry_manufacturer.get()  # Get the manufacturer
    qty_taken = int(entry_qty.get())  # Get the quantity taken
    person_name = entry_name.get()  # Get the person's name
    reason = entry_reason.get()  # Record the reason for picking the item

    # Load the current inventory
    df = load_inventory()

    # Check if the item exists in the inventory based on Product name, Item code, and Manufacturer
    if not df[(df['Product name'] == product_name) & (df['Item code'] == item_code) & (df['Manufacturer'] == manufacturer)].empty:
        # Get current balance and reorder level for the matching product, item code, and manufacturer
        current_balance = df.loc[(df['Product name'] == product_name) & (df['Item code'] == item_code) & (df['Manufacturer'] == manufacturer), 'Balance (Qty)'].values[0]
        reorder_level = df.loc[(df['Product name'] == product_name) & (df['Item code'] == item_code) & (df['Manufacturer'] == manufacturer), 'Reorder level'].values[0]

        # Check if there is enough quantity available
        if current_balance >= qty_taken:
            # Update balance after taking the item
            new_balance = current_balance - qty_taken
            df.loc[(df['Product name'] == product_name) & (df['Item code'] == item_code) & (df['Manufacturer'] == manufacturer), 'Balance (Qty)'] = new_balance
            df.loc[(df['Product name'] == product_name) & (df['Item code'] == item_code) & (df['Manufacturer'] == manufacturer), 'Issued?'] = 'Yes'
            df.loc[(df['Product name'] == product_name) & (df['Item code'] == item_code) & (df['Manufacturer'] == manufacturer), 'Date'] = date
            df.loc[(df['Product name'] == product_name) & (df['Item code'] == item_code) & (df['Manufacturer'] == manufacturer), 'Comments'] = f"Taken by {person_name}, Reason: {reason}"

            # Save the updated inventory
            save_inventory(df)
            messagebox.showinfo("Success", f"{qty_taken} {product_name} taken successfully! Remaining balance: {new_balance}")

            # Check if the new balance is less or equal to the reorder level and trigger reorder if necessary
            if new_balance < reorder_level:
                check_reorder(df, product_name, item_code, manufacturer)

        else:
            messagebox.showerror("Error", f"Not enough {product_name} in stock! Available: {current_balance}, Requested: {qty_taken}")
    else:
        # If no matching product, item code, and manufacturer is found
        messagebox.showerror("Error", "Product not found in inventory or mismatch in item code/manufacturer! Please check and try again.")

# Check if the balance falls below the reorder level
def check_reorder(df, product_name, item_code, manufacturer):
    # Find the balance and reorder level based on Product name, Item code, and Manufacturer
    balance = df.loc[(df['Product name'] == product_name) & (df['Item code'] == item_code) & (df['Manufacturer'] == manufacturer), 'Balance (Qty)'].values[0]
    reorder_level = df.loc[(df['Product name'] == product_name) & (df['Item code'] == item_code) & (df['Manufacturer'] == manufacturer), 'Reorder level'].values[0]

    # Check if the balance is below or equal to the reorder level
    if balance <= reorder_level:
        send_email_alert(product_name, item_code, manufacturer, balance, reorder_level)

# Send an email notification when a product falls below the reorder level
def send_email_alert(product_name, item_code, manufacturer, balance, reorder_level):
    sender_email = "omukutirodney@gmail.com"
    receiver_email = "omukutirodney@gmail.com"
    cc_emails = ["romneyaskolyo@gmail.com", "lukoyedanstone@gmail.com"]  # Add CC emails here
    password = "ltbx ikwe lkmt cqyk"

    # Create subject and body for the email
    subject = f"Reagent '{product_name}' (Code: {item_code}, Manufacturer: {manufacturer}) Inventory Alert"
    body = (f"The balance of the product '{product_name}' (Item Code: {item_code}, Manufacturer: {manufacturer}) "
            f"has fallen below the reorder level.\n"
            f"Current Balance: {balance}\nReorder Level: {reorder_level}\nPlease consider reordering\nThank you.")

    # Create the email message
    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = receiver_email
    msg['Cc'] = ", ".join(cc_emails)  # Add the CC email addresses
    msg['Subject'] = subject
    msg.attach(MIMEText(body, 'plain'))

    # Combine the recipient and CC addresses
    all_recipients = [receiver_email] + cc_emails

    # Send the email (ensure your email setup is correct)
    try:
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(sender_email, password)
        text = msg.as_string()
        server.sendmail(sender_email, all_recipients, text)
        server.quit()
        print(f"Email alert sent successfully to {receiver_email} and CC'd to {', '.join(cc_emails)}.")
    except Exception as e:
        print(f"Failed to send email: {e}")

# Create the main UI
def create_main_ui():
    root = Tk()
    root.title("Lab Supplies Inventory")

    # Introductory message
    intro_label = Label(root, text="Welcome to Dr. Sammy's Lab Inventory", font=("Arial", 30))
    intro_label.grid(row=0, column=0, columnspan=2, pady=10)

    # Display options: Store or Take
    def choose_option(option):
        if option == "store":
            store_window()
        elif option == "take":
            take_window()

    # Create buttons for storing and taking items
    store_button = Button(root, text="Store Reagents/Consumables", command=lambda: choose_option("store"))
    store_button.grid(row=1, column=0, padx=20, pady=10)

    take_button = Button(root, text="Take Reagents/Consumables", command=lambda: choose_option("take"))
    take_button.grid(row=1, column=1, padx=20, pady=10)

    root.mainloop()

# Create UI for storing items
def store_window():
    window = Toplevel()
    window.title("Store Inventory")

    global entry_date, entry_item_code, entry_product_name, entry_manufacturer, entry_received, entry_b_fwd
    global entry_qty, entry_unit, entry_balance_qty, entry_location, entry_reorder_level, entry_comments

    # Labels and entries
    Label(window, text="Date").grid(row=0, column=0)
    entry_date = Entry(window)
    entry_date.grid(row=0, column=1)

    Label(window, text="Item Code").grid(row=1, column=0)
    entry_item_code = Entry(window)
    entry_item_code.grid(row=1, column=1)

    Label(window, text="Product Name").grid(row=2, column=0)
    entry_product_name = Entry(window)
    entry_product_name.grid(row=2, column=1)

    Label(window, text="Manufacturer").grid(row=3, column=0)
    entry_manufacturer = Entry(window)
    entry_manufacturer.grid(row=3, column=1)

    Label(window, text="Received?").grid(row=4, column=0)
    entry_received = Entry(window)
    entry_received.grid(row=4, column=1)

    Label(window, text="B/fwd?").grid(row=5, column=0)
    entry_b_fwd = Entry(window)
    entry_b_fwd.grid(row=5, column=1)

    Label(window, text="Quantity").grid(row=6, column=0)
    entry_qty = Entry(window)
    entry_qty.grid(row=6, column=1)

    Label(window, text="Unit").grid(row=7, column=0)
    entry_unit = Entry(window)
    entry_unit.grid(row=7, column=1)

    Label(window, text="Balance (Qty)").grid(row=8, column=0)
    entry_balance_qty = Entry(window)
    entry_balance_qty.grid(row=8, column=1)

    Label(window, text="Location (in store)").grid(row=9, column=0)
    entry_location = Entry(window)
    entry_location.grid(row=9, column=1)

    Label(window, text="Reorder Level").grid(row=10, column=0)
    entry_reorder_level = Entry(window)
    entry_reorder_level.grid(row=10, column=1)

    Label(window, text="Comments").grid(row=11, column=0)
    entry_comments = Entry(window)
    entry_comments.grid(row=11, column=1)

    # Submit button to store the data
    Button(window, text="Store", command=store_inventory).grid(row=12, column=0, columnspan=2)

# Create UI for taking items
def take_window():
    window = Toplevel()
    window.title("Take Inventory")

    global entry_date, entry_product_name, entry_item_code, entry_manufacturer, entry_qty, entry_name, entry_reason

    # Labels and entries
    Label(window, text="Date").grid(row=0, column=0)
    entry_date = Entry(window)
    entry_date.grid(row=0, column=1)

    Label(window, text="Product Name").grid(row=1, column=0)
    entry_product_name = Entry(window)
    entry_product_name.grid(row=1, column=1)

    Label(window, text="Item Code").grid(row=2, column=0)
    entry_item_code = Entry(window)
    entry_item_code.grid(row=2, column=1)

    Label(window, text="Manufacturer").grid(row=3, column=0)
    entry_manufacturer = Entry(window)
    entry_manufacturer.grid(row=3, column=1)

    Label(window, text="Quantity Taken").grid(row=4, column=0)
    entry_qty = Entry(window)
    entry_qty.grid(row=4, column=1)

    Label(window, text="Person's Name").grid(row=5, column=0)
    entry_name = Entry(window)
    entry_name.grid(row=5, column=1)

    Label(window, text="Reason for Taking").grid(row=6, column=0)
    entry_reason = Entry(window)
    entry_reason.grid(row=6, column=1)

    # Submit button to take the item
    Button(window, text="Take", command=take_inventory).grid(row=7, column=0, columnspan=2)

# Run the main UI
create_main_ui()

