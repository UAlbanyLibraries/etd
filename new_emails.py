import os
import csv
import openpyxl

workspace = "\\\\Lincoln\\Library\\ETDs"
contacts_path = os.path.join(workspace, "MMSIDs", "contacts.csv")
contacts_updated = os.path.join(workspace, "MMSIDs", "contacts_updated.csv")
emails_path = os.path.join(workspace, "ETDWorkspace", "MASTERS_PHD_2008_2022.xlsx")

wb = openpyxl.load_workbook(filename = emails_path)
sheet = wb.active
alumni = []
for alumnus in sheet:
	alumni.append([alumnus[1].value, alumnus[2].value, alumnus[3].value, alumnus[4].value, alumnus[11].value])

contacts_out = []

total = 0
matches = 0
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

		#print (f"looking for {first} {last}")

		match_count = 0
		for alumnus in alumni:
			mailing_name = alumnus[0]
			first_name = alumnus[1]
			last_name = alumnus[2]
			class_year = alumnus[3]
			new_email = alumnus[4]
			if new_email is None:
				new_email = "Empty"

			if name.lower() in mailing_name.lower():
				match_count += 1
				row[10] = new_email
			elif last.lower() in last_name.lower() and first.lower() in first_name.lower():
				match_count += 1
				row[10] = new_email
			elif first.lower() in mailing_name.lower() and last.lower() in mailing_name.lower():
				match_count += 1
				row[10] = new_email

		if match_count > 1:
			match_count = 0
			for alumnus in alumni:
				mailing_name = alumnus[0]
				first_name = alumnus[1]
				last_name = alumnus[2]
				class_year = alumnus[3]
				new_email = alumnus[4]
				if new_email is None:
					new_email = "Empty"

				if last.lower() in last_name.lower() and first.lower() in first_name.lower():
					if year == class_year:
						match_count += 1
						row[10] = new_email
				elif first.lower() in mailing_name.lower() and last.lower() in mailing_name.lower():
					if year == class_year:
						match_count += 1
						row[10] = new_email

		if match_count == 0 and " " in first:
			match_count = 0
			for alumnus in alumni:
				mailing_name = alumnus[0]
				first_name = alumnus[1]
				last_name = alumnus[2]
				class_year = alumnus[3]
				new_email = alumnus[4]
				if new_email is None:
					new_email = "Empty"

				if last.lower() in last_name.lower() and first.split(" ")[0].lower() in first_name.lower():
					if year == class_year:
						match_count += 1
						row[10] = new_email
				elif first.split(" ")[0].lower() + last.lower() in mailing_name.lower():
					if year == class_year:
						match_count += 1
						row[10] = new_email


		if match_count == 1:
			matches += 1
		elif match_count == 0:
			print (f"Couldn't find {first} {last}")
		else:
			print (f"Multiple matches for {first} {last}")

		contacts_out.append(row)


print (f"Found {matches} of {total}")
# Write contacts to CSV
with open(contacts_updated, "w", newline="", encoding="utf-8") as f:
	writer = csv.writer(f)
	writer.writerows(contacts_out)