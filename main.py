from invoices import *
# Set your Stripe API key




# Load all invoices
json_to_excel("output/json_no_metadata.json","output/json_no_metadata.xlsx")
#TODO check if charges are same as invoices 
#TODO check if all currencies are the same
#TODO check payment intents 

#TODO get fields subscription circle and stuff while at it check for ghosts

#Constataion : beacoup de ghosts no metadata