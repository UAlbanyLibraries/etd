import os
import csv
from packages import ETD

if os.name == "nt":
	workspace = "\\\\Lincoln\\Library\\ETDs\\ETDWorkspace"
	storage = "\\\\Lincoln\\Masters\\ETD-storage"
else:
	workspace = "/media/Library/ETDs/ETDWorkspace"
	storage = "/media/Masters/ETD-storage"

matches_file = os.path.join(workspace, "output.csv")
count = 0
with open(matches_file, newline='', encoding="utf-8") as csvfile:
	reader = csv.reader(csvfile)
	for row in reader:
		count += 1
		if count > 1:
			mms_id = row[0]
			etd_id = row[1]
			year = row[3]

			etd_path = os.path.join(storage, year, etd_id)

			etd = ETD()
			etd.load(etd_path)
			print (f"adding {mms_id}")

			etd.bag.info['MMS-ID'] = mms_id
			etd.bag.save()