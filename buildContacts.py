import os
import csv
from tqdm import tqdm
from packages import ETD

etd_root = "\\\\Lincoln\\Masters\\ETD-storage"

contacts_path = "\\\\Lincoln\\Library\\ETDs\\MMSIDs\\contacts.csv"

out_list = [["Opt out",
			"etd_id",
			"name",
			"first_name",
			"last_name",
			"title",
			"complete_date",
			"embargo",
			"email in XML",
			"response",
			"email 2",
			"response_2",
			"email 3",
			"response_3"]]


etd_list = []
for year in os.listdir(etd_root):
	for etd in os.listdir(os.path.join(etd_root, year)):
		etd_list.append([os.path.join(etd_root, year, etd), year])

count = 0
for thing in tqdm(etd_list):
	
	etd_path, year = thing
	count += 1

	if int(year) < 2023:

		etd = ETD()
		etd.load(etd_path)

		if etd.bag.info["Embargo-Active"] == "True":
			embargo_string = etd.bag.info["Embargo-Date"]
		else:
			embargo_string = ""

		etd_out = ["",
		etd.etd_id,
		etd.bag.info["Author"],
		etd.bag.info["First-Name"],
		etd.bag.info["Last-Name"],
		etd.bag.info["Submitted-Title"],
		etd.bag.info["Completion-Date"],
		embargo_string,
		etd.bag.info["ProQuest-Email"],
		"",
		"",
		"",
		"",
		""
		]

		out_list.append(etd_out)

	
# Write contacts to CSV
with open(contacts_path, "w", newline="", encoding="utf-8") as f:
	writer = csv.writer(f)
	writer.writerows(out_list)