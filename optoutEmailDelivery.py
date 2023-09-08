import getopt
import smtplib
import sys

from openpyxl import load_workbook
from email.message import EmailMessage
from email.headerregistry import Address
from email.utils import make_msgid

import envconfig

##import csv
##import requests
##import os
##import shutil
##import sys
##import getopt
##import xlwt
##from packages import ETD
##from lxml import etree


################################################################################
## ETD Optout Email Delivery Script                                           ##
################################################################################
## This Script does the following:                                            ##
##  1.  Reads through a reference input XLSX (provided by Greg)               ##
##  2.  For each entry, send an email to Columns I and K from template        ##
################################################################################

def main(argv):
    print('Initiating...');
    
    #setup
    input_file = ''
    
    # script consumes user-provided values
    # -i flags the input file name
    opts, args = getopt.getopt(argv,'hi:o:',['ifile=','ofile='])
    for opt, arg in opts:
        if opt == '-h':
            print ('migration.py -i <inputfile>')
            sys.exit()
        elif opt in ('-i', '--ifile'):
            input_file = arg
    
    #continue only if input arguments are not empty
    if not input_file:
        print('Input File not set')
    else:
        # read input file
        print('Reading data...')
        
        sheet = load_workbook(filename = input_file, read_only = True)['contacts_updated']
        header = []
        entries = []
        
        for i, row in enumerate(sheet):
            if i == 0:
                header = row
            else:
                entries.append(convert_row(header,row))
        
        for i,entry in enumerate(entries):
            author = entry['name']
            
            emails = [entry['email in XML']]
            
            if len(entry['email 2']) > 0:
                emails.append(entry['email 2'])
            if len(entry['email 3']) > 0:
                emails += entry['email 3'].split(', ')
            
            approved_embargo = len(entry['embargo']) > 0
            
            print('Entry: ' + str(i) + ', Name: ' + author)
            
            send_emails(author, emails, approved_embargo)

#converts a row into an object
def convert_row(headers,row):
    data = {}
    
    for i,header in enumerate(headers):
        value = row[i].value
        
        if value is None:
            data[header.value] = ''
        else:
            data[header.value] = value
    
    return data

#sends emails to addresses listed in each entry
def send_emails(author, emails, approved_embargo):
    smtp_runner = smtplib.SMTP(envconfig.smtp_host, envconfig.smtp_port)
    
    addresses = []
    
    for email in emails:
        print(email)
        parts = email.split('@')
        addresses.append(Address(author, parts[0], parts[1]))
    
    
    message = '''
    <p>Dear {author},</p>
    
    <p>As part of an effort to enhance the web presence of our graduate programs and showcase student work at the University at Albany, we are writing to inform you that in the following months, the University Libraries will begin migrating electronic theses and dissertations (ETDs) from ProQuest, where this work is currently shared, into <a href="http://scholarsarchive.library.albany.edu/">Scholars Archive</a>, the University's open access repository.</p>
    
    <p>Distributing ETDs in Scholars Archive offers our alumni numerous benefits, and including your thesis or dissertation in Scholars Archive will expand the reach of your work, allowing it to be accessed by a global readership.</p>
    
    <p>UAlbany ETDs will still be available and searchable within ProQuest with this initiative, and <strong>you still retain copyright of your thesis or dissertation, allowing you to publish your own work at any time with any publisher.</strong></p>'''.format(author = author, email = email) + ('''<p>[Additionally, if you have a Graduate School-approved embargo, it will be respected in Scholars Archive, as in ProQuest.]</p>''' if approved_embargo else '') + '''<p>If you do not wish to participate in this initiative, please complete this form (<a href="https://albany.libwizard.com/f/retro_etd">https://albany.libwizard.com/f/retro_etd</a>) by {deadline} (one month) indicating that you do not want your ETD added to the repository.</p>
    
    <p>This move to share UAlbany ETDs via Scholars Archive is a valuable expansion on and improvement of the essential role the Libraries and Archives play in preserving and sharing these publications. If you have any questions or concerns, or if you would like to learn about creating an author dashboard within Scholars Archive, then please contact us at <a href="mailto:scholcomm@albany.edu">scholcomm@albany.edu</a>.</p>
    
    <p>Thank you for your time and for continuing to be an invaluable part of the UAlbany community. We are so proud of your work!</p>
    
    <p>Sincerely,</p>
    
    <p>The University at Albany Libraries</p>'''.format(deadline = 'October 8, 2023')
    
    msg = EmailMessage()
    msg['Subject'] = 'test'
    msg['From'] = Address('University at Albany Libraries', 'libwww', 'albany.edu')
    msg['To'] = addresses
    msg.set_content(message, subtype="html")
    
    smtp_runner.send_message(msg)










#pulls bib data using Alma API
#returns bib JSON or None
def get_bib(mmsid):
    url = 'https://api-na.hosted.exlibrisgroup.com/almaws/v1/bibs/' + mmsid + '?apikey=' + envconfig.api_key
    headers = {'Accept':'application/json'}
    
    response = requests.get(url, headers=headers)
    
    if response.status_code == 200:
        return response.json()
    
    return None

#loads the associated etd object from storage
#returns ETD object or None
def get_etd(etd_id,xml_id,year):
    etd_record = ETD()
    
    if not os.path.isdir(os.path.join(envconfig.storage_directory, year, etd_id)):
        return None;
    else:
        etd_record.load(os.path.join(envconfig.storage_directory, year, etd_id))
        return etd_record

#copies the etd pdf file into a receiving directory visible to BePress
#returns a url string for said file as per the directory's mounting on the Apps Server or empty string
def deposit_pdf(etd_record):
    shutil.copyfile(etd_record.pdf_file,os.path.join(envconfig.working_directory, 'SA_uploads', etd_record.etd_id + '.pdf'))
    
    return 'https://apps.library.albany.edu/sa_uploads/' + etd_record.etd_id + '.pdf'

#trigger for main method
if __name__ == "__main__":
    main(sys.argv[1:])