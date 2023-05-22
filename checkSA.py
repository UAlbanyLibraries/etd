import os
import re
import csv
import json
import time
import requests
import openpyxl
import urllib.parse
from tqdm import tqdm
from fuzzywuzzy import fuzz

# if windows
if os.name == "nt":
    catalog_root = "\\\\Lincoln\\Library\\ETDs\\MMSIDs"
else:
    catalog_root = "/media/Library/ESPYderivatives/etds_testing"
pre2008 = os.path.join(catalog_root, "MMSID_1967_2007.xlsx")
post2008 = os.path.join(catalog_root, "MMSID_2008_2022.xlsx")

output_path = os.path.join(catalog_root, "output.csv")

#define a function to check if a json result from the bepress API is actually a match
def isMatch(result, title, author, date):
    #print ("Comparing:")
    #print (f"\t{title}, date")
    #print (f"\t{author}")
    #print ("To:")
    #print (f"\t{result['title']}, {result['publication_date'][:4]}")

    stipped_title = re.sub(r'[^a-zA-Z]', ' ', title.lower().strip()).strip()
    stipped_author = re.sub(r'[^a-zA-Z]', ' ', author.lower().strip()).strip()
    stipped_result_title = re.sub(r'[^a-zA-Z]', ' ', result["title"].lower().strip()).strip()

    if stipped_title == stipped_result_title:
        return True
    else:
        title_ratio = fuzz.ratio(stipped_title, stipped_result_title)
        author_ratio = 0
        if "author" in result.keys():
            for result_author in result["author"]:
                stipped_result_author = re.sub(r'[^a-zA-Z]', ' ', result_author.lower().strip()).strip()
                auth_ratio = fuzz.ratio(stipped_author, stipped_result_author)
                if auth_ratio > author_ratio:
                    author_ratio = auth_ratio
        if title_ratio > 80 and author_ratio > 80:
            return True
        elif title_ratio > 60 and author_ratio > 60 and result["publication_date"].startswith(date):
            return True
        else:
            return False

# bepress often 504s, wonderful
# It also just disconnects without sending a response so we have to try/except
# this fuction manages bepress api calls and tries again after 500 or 504 errors
def bepress(query, headers):
    try_count = 0
    while try_count < 10:
        try:
            r = requests.get(query, headers=headers)
            if r.status_code == 200:
                return r
            else:
                print (r.text)
                if str(r.status_code).startswith("5"):
                    try_count += 1
                    print (f"Waiting before try {str(try_count)}...")
                    time.sleep(10)
                    # Will continue the while loop
                else:
                    raise ValueError(f"ERROR: bepress returned {str(r.status_code)}")
        except:
            print ("Dropped connection. Waiting and then trying again...")
            time.sleep(10)
            try_count += 1

    raise ValueError(f"ERROR: Number of tries exceeded.")


# Build the catalog data into a list in memory so we can find stuff in it later
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

            if row[0].value.endswith(".)"):
                date = row[0].value[-6:-2]
            else:
                date = row[0].value[-5:-1]
            if len (date) != 4:
                print ("ERROR: " + date) 
            try:
                dateNumber = int(date)
                if dateNumber > 2023 or dateNumber < 2008:
                    print ("ERROR: " + date)
            except:
                print ("ERROR: " + row[0].value)

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
            records.append([title, author, mms_id, date])

# Set up the bepress API calls
token_file = open("token.txt", "r")
token = token_file.read().strip()
token_file.close()
bepressURL = "https://content-out.bepress.com/v2/scholarsarchive.library.albany.edu/query"
headers = {"Authorization": token}

# Start CSV output data with headers
sa_matches_header = ["mms_id", "zip_id", "title", "author", "date", "sa_url"]
sa_match_count = 0
with open(os.path.join(catalog_root, "sa_matches.csv"), "w", newline="") as f:
    writer = csv.writer(f)
    writer.writerow(sa_matches_header)
    f.close()

# Read the output file which has mss_ids and matching pdf ids
with open(output_path, "r") as csv_file:
    csv_reader = csv.reader(csv_file, delimiter=',')
    line_count = 0
    for row in tqdm(csv_reader):
        line_count += 1
        if line_count == 1:
            pass
        else:
            #find it in catalog data
            for record in records:
                if row[0] == record[2]:
                    title = record[0]
                    author = record[1]
                    date = record[3]
                    #print (title)
                    #print (author)
                    #print (date)

                    escaped_title = urllib.parse.quote(re.sub(r'[^a-zA-Z]', ' ', title)).strip()
                    escaped_author = urllib.parse.quote(re.sub(r'[^a-zA-Z]', ' ', author)).strip()

                    # query bepress api
                    query = f"{bepressURL}?q={escaped_title} {escaped_author}"
                    #print (query)
                    #print (escaped_title)

                    # query bepress api
                    r = bepress(query, headers)

                    match = False
                    hits = r.json()["query_meta"]["total_hits"]
                    print (f"Found {hits} results.")

                    matchURL = ""
                    matches = 0
                    for result in r.json()["results"]:
                        #print (json.dumps(result, indent=4))
                        if isMatch(result, title, author, date):
                            matches += 1
                            matchURL = result["url"]

                    if matches == 1:
                        print ("found match!")
                        sa_match_count += 1
                        match_row = [row[0], row[1], title, author, date, matchURL]
                    else:
                        match_row = [row[0], row[1], title, author, date, ""]

                    with open(os.path.join(catalog_root, "sa_matches.csv"), "a", newline="", encoding="utf8") as f:
                        writer = csv.writer(f)
                        writer.writerow(match_row)

    print (f"Of {line_count - 1} ETDs, found {sa_match_count} matches in Scholar's Archive")

