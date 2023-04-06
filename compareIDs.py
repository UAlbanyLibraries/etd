import os
import re
import csv
import openpyxl
from tqdm import tqdm
from lxml import etree
from fuzzywuzzy import fuzz

print ("hello world")

# Lets get a big list of titles and mms_ids
# \ is an escape so you have to double them up for windows paths
catalog_root = "\\\\Lincoln\\Library\\ETDs\\MMSIDs"
pre2008 = os.path.join(catalog_root, "MMSID_1967_2007.xlsx")
post2008 = os.path.join(catalog_root, "MMSID_2008_2022.xlsx")

# empty dict to add catalog data too
records = []
print ("reading catalog export...")
row_count = 0
wb = openpyxl.load_workbook(filename=post2008, read_only=True)
for sheet in wb.worksheets:
	for row in sheet.rows:
		row_count += 1
		#skip header
		if row_count > 1:
			title_text = row[3].value
			mms_id = row[24].value

			# why is there other junk in the title field?
			# This gets all the text before " / ". Dunno if this is safe
			if " / by " in title_text:
				title, author = title_text.split(" / by ")
			elif " / " in title_text:
				title, author = title_text.split(" / ")
			else:
				#print (title_text)
				title, author = title_text.split(" by ")

			# add to records as a list
			records.append([title, author, mms_id])



# Now lets loop though all the ETD packages we got from ProQuest
etd_root = "\\\\Lincoln\\Library\\ETDs\\Unzipped"

# Start of list of 
output = [["mms_id", "xml_id"]]

# Start a list of contact emails to write to a .csv at the end
contacts = [["Name", "Current Email", "Future Email"]]

print ("reading ETD packages...")
xml_list = []
for thing in os.listdir(etd_root):
	# exclude supplemental materials directories and pdfs themselves
	if thing.lower().endswith(".xml"):
		xml_list.append(thing)

count = 0
found = 0
missing = 0
missing_list = []

# Loop though the XML
for etd in tqdm(xml_list):
	# get the full path
	xml_file = os.path.join(etd_root, etd)
	xml_id = etd.split("_DATA")[0]

	count += 1

	#if count == 999:
	if count > 0:
		# Read the XML file
		tree = etree.parse(xml_file)
		# Get the root <DISS_submission> element
		root = tree.getroot()

		author_count = 0
		for author in root.find("DISS_authorship"):
			author_count += 1
			# Thay all have one author... for now
			if author_count != 1:
				raise ValueError(f"{etd} has multiple authors")

			name_count = 0
			fullname = []
			for name in author:
				if name.tag == "DISS_name":
					name_count += 1
					if name_count != 1:
						raise ValueError(f"{etd} has multiple names")

					# safely build the name
					if name.find("DISS_fname").text:
						fullname.append(name.find("DISS_fname").text.title())
					if name.find("DISS_middle").text:
						fullname.append(name.find("DISS_middle").text.title())
					if name.find("DISS_surname").text:
						fullname.append(name.find("DISS_surname").text.title())
					if name.find("DISS_suffix").text:
						fullname.append(name.find("DISS_suffix").text.title())
					#print (" ".join(fullname))

			current_email = author.xpath("//DISS_contact[@type='current']/DISS_email")
			future_email = author.xpath("//DISS_contact[@type='future']/DISS_email")
			if len(current_email) > 1 or len(future_email) > 1:
				raise ValueError(f"{etd} has weird contact info")
			else:
				if len(current_email) == 0:
					current_email_text = ""
				else:
					current_email_text = current_email[0].text
				if len(future_email) == 0:
					future_email_text = ""
				else:
					future_email_text = future_email[0].text
				# Make a list with the fullname (as a string), current and future emails
				contact_row = [" ".join(fullname), current_email_text, future_email_text]
				# add it to the CSV as a row
				contacts.append(contact_row)

		# Get the title
		title = root.xpath("//DISS_description/DISS_title")
		if len(title) != 1:
			raise ValueError(f"{etd} has multiple titles")
		title_text = title[0].text

		match_count = 0
		# First look for easy matches
		for record in records:
			if title_text.lower() in record[0].lower():
				match_count += 1
				catalog_id = record[2]
				# end loop
				break
			else:
				# then try removing punctuation and tokenizing
				xml_tokens = re.sub(r'[^a-zA-Z0-9]\s', ' ', title_text.lower()).split()
				catalog_tokens = re.sub(r'[^a-zA-Z0-9]\s', ' ', record[0].lower()).split()

				if " ".join(xml_tokens) in " ".join(catalog_tokens):
					match_count += 1
					catalog_id = record[2]


		# for others use fuzzy matching
		if match_count == 0:
			for record in records:
				compare_title = fuzz.ratio(title_text, record[0])
				compare_authors = fuzz.ratio(fullname, record[1])
				# 80 is probably good
				if compare_title > 80 or compare_authors > 80:
					match_count += 1
					catalog_id = record[2]
				elif compare_title > 50 and compare_authors > 50:
					match_count += 1
					catalog_id = record[2]
			if match_count == 0:
				# if still no matches, try compating the stripped lowercase
				for record in records:
					stipped_xml = re.sub(r'[^a-zA-Z0-9]\s', ' ', title_text.lower())
					stipped_catalog = re.sub(r'[^a-zA-Z0-9]\s', ' ', record[0].lower())
					compare_stripped = fuzz.ratio(stipped_xml, stipped_catalog)
					if compare_stripped > 80:
						match_count += 1
						catalog_id = record[2]



		# make sure there is only one match
		if match_count == 1:
			found += 1
			output.append([catalog_id, xml_id])
		else:
			if match_count > 1:
				print (f"multiple {str(match_count)} matches --> {xml_id}")
			missing += 1
			missing_list.append(title_text)

print (missing_list)
print (f"Found {str(found)} of {str(count)} ETDs. {str(missing)} missing.")		

# Write contacts to CSV
output_path = os.path.join(catalog_root, "output.csv")
with open(output_path, "w", newline="") as f:
	writer = csv.writer(f)
	writer.writerows(output)

# Write contacts to CSV
contacts_path = os.path.join(catalog_root, "contacts.csv")
with open(contacts_path, "w", newline="") as f:
	writer = csv.writer(f)
	writer.writerows(contacts)
