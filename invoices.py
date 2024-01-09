from openpyxl import Workbook
import stripe
from datetime import datetime
import time
from app_secrets import *
import json
import os

stripe.api_key = your_stripe_api_key
def json_to_excel(json_filename, excel_filename):
    """
    Converts a JSON file to an Excel file.
    """
    with open(json_filename, 'r') as file:
        data = json.load(file)

    workbook = Workbook()
    sheet = workbook.active

    # Headers based on the Stripe invoice object
    headers = ['ID', 'Customer ID', 'Amount Paid', 'Currency', 'Status', 'Invoice Date']
    sheet.append(headers)

    for invoice in data:
        invoice_date = datetime.fromtimestamp(invoice['created']).strftime('%Y-%m-%d')
        row = [
            invoice['id'],
            invoice.get('customer', ''),
            invoice.get('amount_paid', 0),
            invoice.get('currency', ''),
            invoice.get('status', ''),
            invoice_date
        ]
        sheet.append(row)

    workbook.save(excel_filename)
def fetch_all_invoices_and_save_to_json(filename):
    """
    Fetches all invoices from Stripe API and saves them to a JSON file.
    """
    invoices = stripe.Invoice.list(limit=100)  # Adjust limit as needed
    invoices_json = []

    for invoice in invoices.auto_paging_iter():
        invoices_json.append(invoice.to_dict())

    with open(filename, 'w') as file:
        json.dump(invoices_json, file, indent=4)
def find_metadata_keys(data, parent_key='', unique_keys=set()):
    """
    Recursively finds keys that contain 'metadata' in a nested dictionary.
    """
    if isinstance(data, dict):
        for key, value in data.items():
            current_key = f"{parent_key}.{key}" if parent_key else key
            if 'metadata' in key:
                unique_keys.add(current_key)
            if isinstance(value, dict):
                find_metadata_keys(value, current_key, unique_keys)
    return unique_keys

def extract_metadata_keys_from_json(filename):
    """
    Extracts unique metadata keys from JSON file.
    """
    with open(filename, 'r') as file:
        invoices = json.load(file)

    unique_metadata_keys = set()
    for invoice in invoices:
        find_metadata_keys(invoice, unique_keys=unique_metadata_keys)

    return unique_metadata_keys


