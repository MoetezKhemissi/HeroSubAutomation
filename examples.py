'''
# Fetch all invoices and write to Excel
invoices = fetch_all_invoices()
invoices_to_excel(invoices, 'all_invoices.xlsx')
'''

'''
# Fetch n invoices
n = 10  # Number of recent invoices you want to fetch
recent_invoices = fetch_last_n_invoices(n)
for invoice in recent_invoices.auto_paging_iter():
    # Process each invoice as needed
    print(invoice)
    
    '''