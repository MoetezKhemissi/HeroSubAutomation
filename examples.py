'''
# Fetch all invoices and write to Excel
invoices = fetch_all_invoices()
invoices_to_excel(invoices, 'all_invoices.xlsx')
'''

'''
# # Load all invoices
all_invoices = load_invoices_from_json('output/all_invoices.json')

# Extract invoices with no metadata
filter_and_save_invoices(all_invoices, lambda inv: is_metadata_empty(inv.get('metadata')), 'output/json1.json')

# Extract invoices with no subscription_details.metadata
filter_and_save_invoices(all_invoices, lambda inv: is_metadata_empty(inv.get('subscription_details', {}).get('metadata')), 'output/json2.json')
    
    '''

# Metadata keys are {'metadata', 'subscription_details.metadata'}