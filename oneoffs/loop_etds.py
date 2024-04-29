sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
from packages import etd_loop, etd_total
from tqdm import tqdm
import csv

csv_out = [["Year", "ID", "Author", "Title", "Embargo-Date", "MMS-ID", "Supplemental-Path"]]

print (etd_total())
count = 0
for etd in tqdm(etd_loop(), total=etd_total()):

	if etd.supplemental:
		count += 1
		item = []
		item.append(etd.year)
		item.append(etd.etd_id)
		item.append(etd.bag.info["Author"])
		item.append(etd.bag.info["Submitted-Title"])
		item.append(etd.bag.info["Embargo-Date"])
		if "MMS-ID" in etd.bag.info.keys():
			item.append(etd.bag.info["MMS-ID"])
		else:
			item.append("")
		item.append(etd.bag.info["Supplemental-Path"])

		csv_out.append(item)

print (count)
with open('supplemental.csv', 'w', encoding="cp1252", newline='') as f:
    writer = csv.writer(f)
    writer.writerows(csv_out)