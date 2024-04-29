import os
import sys
from openpyxl import load_workbook

sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
from packages import ETD

root_path = "\\\\Lincoln\\Library\\ETDs"
etd_storage = "\\\\Lincoln\\Masters\\ETD-storage"

#old_embargos_path = os.path.join(root_path, "embargos.xlsx")
new_embargos_path = os.path.join(root_path, "ETD Embargoed Submissions 2023 to 4-10-24.xlsx")

wb = load_workbook(new_embargos_path)
sheet = wb.active

row_count = 0
for row in sheet:
	row_count += 1
	if row_count > 1:
		year = row[6].value.split("-")[0]
		lname = row[3].value
		fname = row[4].value
		title = row[10].value
		embargo_date = row[21].value

		if not year == "2024":
			year_folder = os.path.join(etd_storage, year)

			found = 0
			etd_path = ""
			for package in os.listdir(year_folder):
				if package.lower().startswith(lname.lower().replace(" ", "_")):
					found += 1
					
					etd_path = os.path.join(year_folder, package)
					"""
					etd = ETD()
					etd.load(etd_path)
					if lname.lower() in etd.bag.info["Last-Name"].lower():
						if fname.lower() in etd.bag.info["First-Name"].lower():
							found += 1
					"""

			if found == 1:
				etd = ETD()
				etd.load(etd_path)
				if etd.bag.info["Embargo-Date"] != embargo_date:
					print (f"Error: {etd.etd_id} listed {etd.bag.info['Embargo-Date']} does not match {embargo_date}")
					print ("\t" + title)

			"""
			if found == 0:
				print (f"Cannot find {lname} in {year}")
			elif found > 1:
				print (f"Found {found} matches for {lname} in {year}")
			"""
