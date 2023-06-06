import os
import csv
import openpyxl

workspace = "\\\\Lincoln\\Library\\ETDs"
contacts_path = os.path.join(workspace, "MMSIDs", "contacts.csv")
emails_path = os.path.join(workspace, "ETDWorkspace", "MASTERS_PHD_2008_2022.xlsx")

wb = openpyxl.load_workbook(filename = emails_path)
sheet = wb.active
alumni = []
for alumnus in sheet:
	alumni.append([alumnus[1].value, alumnus[2].value, alumnus[3].value, alumnus[4].value, alumnus[11].value])

print ("here")
total = 0
with open(contacts_path, newline='', encoding="utf-8") as csvfile:
	reader = csv.reader(csvfile)
	for row in reader:
		total += 1
		etd_id = row[1]
		name = row[2]
		first = row[3]
		last = row[4]
		year = row[6]
		old_email = row[8]

		match_count = 0
		for alumnus in alumni:
			mailing_name = alumnus[0]
			first_name = alumnus[1]
			last_name = alumnus[2]
			class_year = alumnus[3]
			new_email = alumnus[4]

			if last.lower() in last_name.lower() and first.lower() in first_name.lower():
				match_count += 1

print (f"Found {match_count} of {total}")