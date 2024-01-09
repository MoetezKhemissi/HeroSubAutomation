from invoices import *
# Set your Stripe API key



os.makedirs('output', exist_ok=True)
fetch_all_invoices_and_save_to_json('output/all_invoices.json')
json_to_excel('output/all_invoices.json', 'output/all_invoices.xlsx')




#TODO check if charges are same as invoices 
#TODO check if all currencies are the same
#TODO check payment intents 

#TODO get fields subscription circle and stuff while at it check for ghosts
