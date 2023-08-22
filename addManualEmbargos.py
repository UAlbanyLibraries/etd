from packages import ETD
import os

if os.name == "nt":
	root = "\\\\Lincoln\\Masters\\ETD-storage"
else:
	root = "/media/Masters/ETD-storage"

# this script adds manual embargos approved by the Graduate School after submission to the preservation ETD packages

manual_embargos = [
	["Kristen", "Kaytes", "2025-07-26"],
	["Marissa", "Louis", "2025-05-22"],
	["Rachel", "Netzband", "2025-05-22"],
	["Megan", "Chambers", "2025-02-28"],
	["Meghan", "Appley", "2023-12-21"],
	["Hadi", "Habibzadeh", "2025-01-27"],
	["Nathan", "Bartlett", "2024-05-18"]
]


for year in os.listdir(root):
	if int(year) > 2020:
		print ("looking in " + year)
		for etd_folder in os.listdir(os.path.join(root, year)):
			etd_path = os.path.join(root, year, etd_folder)

			etd = ETD()
			etd.load(etd_path)
			modified = False

			for embargo in manual_embargos:
				if etd.bag.info["First-Name"].lower().strip() == embargo[0].lower().strip():
					if etd.bag.info["Last-Name"].lower().strip() == embargo[1].lower().strip():
						print (f"found match for {embargo[0]} {embargo[1]} --> {year}, {etd_folder}")

						#print (embargo[2])
						etd.bag.info["Embargo"] = "True"
						etd.bag.info["Embargo-Active"] = "True"
						etd.bag.info["Embargo-Date"] = embargo[2]
						etd.bag.info["Embargo-Type"] = "Specified Date"
						etd.bag.info["Embargo-Note"] = "Manual embargo approved by Graduate School"
						etd.bag.save()
						modified = True

			if modified:
				etd = ETD()
				etd.load(etd_path)
				print (etd.bag.is_valid())