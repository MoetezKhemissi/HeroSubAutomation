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

def extract_sub_keys(data, target_keys):
    """
    Extracts sub-keys from specified metadata fields in a dictionary.
    """
    sub_keys = set()
    for target_key in target_keys:
        keys = target_key.split('.')
        value = data
        for key in keys:
            value = value.get(key, {})
            if not isinstance(value, dict):
                break
        if isinstance(value, dict):
            sub_keys.update(value.keys())
    return sub_keys
def extract_metadata_sub_keys_from_json(filename, metadata_keys):
    """
    Extracts unique sub-keys from metadata fields in JSON file.
    """
    with open(filename, 'r') as file:
        invoices = json.load(file)

    unique_sub_keys = set()
    for invoice in invoices:
        sub_keys = extract_sub_keys(invoice, metadata_keys)
        unique_sub_keys.update(sub_keys)

    return unique_sub_keys

def load_invoices_from_json(filename):
    """
    Load invoices from a JSON file.
    """
    with open(filename, 'r') as file:
        return json.load(file)

def is_metadata_empty(metadata):
    """
    Check if metadata is empty or doesn't exist.
    """
    return not metadata or len(metadata) == 0

def filter_and_save_invoices(invoices, condition_check, output_filename):
    """
    Filters invoices based on a condition and saves them to a JSON file.
    """
    filtered_invoices = [inv for inv in invoices if condition_check(inv)]
    with open(output_filename, 'w') as file:
        json.dump(filtered_invoices, file, indent=4)


def filter_and_save_invoices(invoices, output_filename):
    """
    Filters invoices based on the condition of having no metadata and no subscription_details.metadata, then saves them to a JSON file.
    """
    filtered_invoices = [
        inv for inv in invoices
        if is_metadata_empty(inv.get('metadata')) and is_metadata_empty(inv.get('subscription_details', {}).get('metadata'))
    ]
    with open(output_filename, 'w') as file:
        json.dump(filtered_invoices, file, indent=4)