# Lab Inventory System

This Python-based lab inventory management system has a graphical user interface (GUI) built using Tkinter. It allows users to store, take, and manage reagents and consumables.

## Features
- Store reagents/consumables
- Track usage and quantities
- Reorder alerts via email when supplies are low
- Google Sheets integration for real-time data storage

## Requirements
- Python 3.x
- Tkinter
- Pandas
- Gspread
- OAuth2Client
- smtplib (for email alerts)

## Setup Instructions

1. Clone the repository:
   ```bash
   git clone https://github.com/Rodneyomukuti/lab_inventory_system.git
   ```

2. Install the required packages:

   ```bash
   pip install pandas gspread oauth2client tkinter
   ```

3. Set up Google Sheets API credentials:

Create a `credentials.json` file in the project directory and follow [this guide](https://docs.gspread.org/en/latest/oauth2.html) to set up OAuth2 credentials for Google Sheets API.

4. Update the `inventory-438319-ae30963a455c.json` file with your credentials to access the Google Sheets.

5. Run the application:

   ```bash
   python lab_inventory.py
   ```

6. Interact with the inventory through the GUI


