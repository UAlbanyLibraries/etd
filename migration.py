import csv
import requests
import os
import xlwt
from lxml import etree

def main(argv):
    print('Initiating...');
    
    #setup
    root_directory = '\\\\Lincoln\\Library\\ETDs'
    input_file = ''
    output_file = ''
    
    # script consumes user-provided values
    # -i flags the input file name
    # -o flags the output file name
    opts, args = getopt.getopt(argv,'hi:o:',['ifile=','ofile='])
    for opt, arg in opts:
        if opt == '-h':
            print ('migration.py -i <inputfile> -o <outputfile>')
            sys.exit()
        elif opt in ('-i', '--ifile'):
            input_file = arg
        elif opt in ('-o', '--ofile'):
            output_file = arg
    
    #continue only if input and output arguments are not empty
    if not input_file:
        print('Input File not set')
    elif not output_file:
        print('Output File not set')
    else
        print('Gathering data...')
        
        #create output workbook
        output_wb = xlwt.Workbook()
        output_ws = output_wb.add_sheet(' ')
        
        # read input file
        with open(input_file, newline='') as input_csv:
            input_reader = csv.reader(input_csv)
            input_headers = next(input_reader)
            
            for row in input_reader:
                mmsid = row_array[0]
                dataxmlid = row_array[1]
                
                # retrieve bib record with array[0] value
                bib_record = get_bib(mmsid)
                
                # retrieve ..._DATA.xml file with array[1] value
                data_record = get_dataxml(dataxmlid)
                
                # write entry into output_ws
            
            
        #save output file
        output_wb.save(output_file)

def get_bib(mmsid):
    url = 'https://api-na.hosted.exlibrisgroup.com/almaws/v1/bibs/' + mmsid + '?apikey=l8xxfcd6b610fe4647e58bc95edc4a629ca0'
    headers = {'Accept':'application/json'}
    
    response = requests.post(url, headers=headers);
    if response.status_code == 200:
        return response.json()
    
    return None

def get_dataxml(name):
    xml_file_path = os.path.join(root_directory, 'Unzipped', name + '_DATA.xml')
    
    with open(xml_file_path) as xml_file:
        xml_reader = xml_file.read()
    return something

if __name__ == "__main__":
    main(sys.argv[1:])