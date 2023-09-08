import os
import openpyxl

root = "\\\\Lincoln\\Library\\ETDs\\MMSIDs"

contacts = os.path.join(root, "contacts.xlsx")
newContacts = os.path.join(root, "Copy of remaining_emails.xlsx")
contactsOutput = os.path.join(root, "contacts_new.xlsx")


print ("reading new contacts...")
row_count = 0
wb = openpyxl.load_workbook(filename=newContacts, read_only=True)
contactsFile = openpyxl.load_workbook(filename=contacts, read_only=False)
contactsSheet = contactsFile.active
for sheet in wb.worksheets:
    for row in sheet.rows:
        row_count += 1
        if row_count > 1:
        	xml_email = row[7].value.strip().lower()
        	new_email = row[8].value

        	print (f"Looking for {xml_email}...")

        	for contactsRow in contactsSheet:
        		if contactsRow[8].value is None:
        			pass
        		else:
        			if contactsRow[8].value.strip().lower() == xml_email:
        				print ("\tfound.")
        				if new_email is None or new_email.strip().lower() == "only ualbany":
        					pass
        				else:
        					if contactsRow[10].value:
        						if not new_email.strip().lower() == contactsRow[10].value.strip().lower():
        							contactsRow[12].value = new_email.strip()

print ("Saving...")
contactsFile.save(contactsOutput)