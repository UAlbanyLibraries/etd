import csv
import requests
import os
import sys
import getopt
import xlwt
from etd import etd
from lxml import etree
import envconfig

working_directory = '\\\\Lincoln\\Library\\ETDs'
storage_directory = '\\\\Lincoln\\Masters\\ETD-storage'

def main(argv):
    print('Initiating...');
    
    #setup
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
    else:
        print('Gathering data...')
        
        #create output workbook
        output_wb = xlwt.Workbook()
        output_ws = output_wb.add_sheet(' ')
        
        
        headers = [
            {'input':'title','output':'title'},
            {'input':'url','output':'fulltext_url'}, #?????
            {'input':'keyword','output':'keywords'},
            {'input':'abstract','output':'abstract'},
            {'input':'First-Name','output':'author1_fname'},
            {'input':'Middle-Name','output':'author1_mname'},
            {'input':'Last-Name','output':'author1_lname'},
            {'input':'Author-Suffix','output':'author1_suffix'},
            {'input':'ProQuest-Email','output':'author1_email'},
            {'input':'inst_name','output':'author1_institution'},
            {'input':'advisor','output':'advisor1'},
            {'input':'advisor2','output':'advisor2'}, #?????
            {'input':'advisor3','output':'advisor3'}, #?????
            {'input':'mms_id','output':'alma_mms_id'},
            {'input':'categorization','output':'disciplines'},
            {'input':'050','output':'call_number'},
            {'input':'',    'output':'carrier_type'},
            {'input':'',    'output':'comments'},
            {'input':'',    'output':'committee_members'},
            {'input':'',    'output':'content_type'},
            {'input':'',    'output':'degree_name'},
            {'input':'',    'output':'department'},
            {'input':'',    'output':'distribution_license'},
            {'input':'060','output':'document_type'},
            {'input':'',    'output':'doi'},
            {'input':'',    'output':'embargo_date'},
            {'input':'',    'output':'genre_form'},
            {'input':'',    'output':'language'},
            {'input':'',    'output':'media_type'},
            {'input':'',    'output':'oa_licenses'},
            {'input':'',    'output':'orcid'},
            {'input':'300','output':'phys_desc'},
            {'input':'date_of_publication','output':'publication_date'},
            {'input':'',    'output':'season'},
            {'input':'',    'output':'rights_statements'},
            {'input':'',    'output':'subjects'}
            ]
        
        #worksheet header row
        for column,header in enumerate(headers):
            output_ws.write(0,column,header['output']);
        
        # read input file
        with open(input_file, newline='', encoding="utf8") as input_csv:
            input_reader = csv.reader(input_csv)
            input_headers = next(input_reader)
            
            for row_count,row_array in enumerate(input_reader):
                mms_id = row_array[0]
                etd_id = row_array[1]
                xml_id = row_array[2]
                year   = row_array[3]
                
                print(' '.join([mms_id, etd_id, year]))
                
                # retrieve records and flatten
                bib_record = get_bib(mms_id)
                etd_record = get_etd(etd_id, xml_id, year)
                
                data = flatten_data(headers,bib_record,etd_record, mms_id,etd_id,xml_id,year)
                
                # write entry into output_ws
                for column,header in enumerate(headers):
                    #print('Column ' + str(column) + ':' + header['output'])
                    output_ws.write(row_count+1,column,data[header['output']])
        #save output file
        output_wb.save(output_file)

#combines data from bib record and etd record to produce a flat object that contains the necessary field values
def flatten_data(headers,bib_record,etd_record, mms_id,etd_id,xml_id,year):
    data = {}
    
    #this fills the row with an error report if the associated input data is corrupt
    if etd_record is None or bib_record is None:
        for i,header in enumerate(headers):
            if i == 0:
                if bib_record == None:
                    data[header['output']] = 'Bib Record Not Retrieved: ' + mms_id
                else:
                    data[header['output']] = 'Bib Record Fine: ' + mms_id
            elif i == 1:
                if etd_record == None:
                    data[header['output']] = 'ETD Record Not Retrieved: ' + etd_id + ' ' + year + ' ' + xml_id
                else:
                    data[header['output']] = 'ETD Record Fine: ' + etd_id + ' ' + year + ' ' + xml_id
            else:
                data[header['output']] = ''
    else:
        bib_xml = etree.XML(bib_record['anies'][0].encode(encoding="UTF-16"))
        etd_xml = etree.parse(etd_record.xml_file)
    
        for header in headers:
            header_input = header['input']
            header_output = header['output']
            
            match header_input:
                case 'title' | 'mms_id' | 'date_of_publication': #taken from bib data json
                    data[header['output']] = bib_record[header['input']]
                case '050' | '060' | '300': #taken from bib data xml
                    datafields = bib_xml.xpath("//datafield[@tag=" + header['input'] +"]")
                    data[header['output']] = ''
                    
                    for i,datafield in enumerate(datafields):
                        if i > 0:
                            data[header['output']] += '\n'
                        subfields = datafield.findall('subfield')
                        
                        for j,subfield in enumerate(subfields):
                            if j > 0:
                                data[header['output']] += ' '
                            data[header['output']] += subfield.text
                case 'abstract' | 'keyword' | 'inst_name': #taken from etd's ..._DATA.xml
                    found_els = etd_xml.find('DISS_' + header['input'])
                    if found_els != None:
                        data[header['output']] = ''.join(found_els.itertext())
                    else:
                        data[header['output']] = ''
                case 'First-Name' | 'Middle-Name' | 'Last-Name' | 'ProQuest-Email' : #taken from etd metadata in bag-info.txt
                    if header['input'] in etd_record.bag.info:
                        data[header['output']] = etd_record.bag.info[header['input']]
                    else:
                        data[header['output']] = ''
                case 'advisor': #taken from etd ..._DATA.xml, specific to advisor group
                    advisors = etd_xml.findall('DISS_advisor')
                    
                    data['advisor1'] = ''
                    data['advisor2'] = ''
                    data['advisor3'] = ''
                    
                    for i,advisor in enumerate(advisors):
                        advisor_name = [];
                        if advisor.find('DISS_fname').text:
                            advisor_name[0] = advisor.find('DISS_fname').text.title()
                        if advisor.find('DISS_middle').text:
                            advisor_name[1] = advisor.find('DISS_middle').text.title()
                        if advisor.find('DISS_surname').text:
                            advisor_name[1] = advisor.find('DISS_surname').text.title()
                        
                        data['advisor' + (i+1)] = ' '.join(advisor_name)
                case 'categorization': #taken from etd ..._DATA.xml, specific to categories group
                    categories = etd_xml.findall('DISS_category')
                    disciplines = ''
                    
                    for category in categories:
                        disciplines = disciplines.join(category.find('DISS_cat_desc')).text
                    data[header['output']] = disciplines
                case _: #unassigned elsewhere, leave blank for now
                    data[header['output']] = ''
    
    return data
    
def get_bib(mmsid):
    url = 'https://api-na.hosted.exlibrisgroup.com/almaws/v1/bibs/' + mmsid + '?apikey=' + envconfig.api_key
    headers = {'Accept':'application/json'}
    
    response = requests.get(url, headers=headers)
    
    if response.status_code == 200:
        return response.json()
    
    return None

def get_etd(etd_id,xml_id,year):
    etd_record = etd()
    
    if not os.path.isdir(os.path.join(storage_directory, year, etd_id)):
        return None;
    else:
        etd_record.load(os.path.join(storage_directory, year, etd_id))
        return etd_record

if __name__ == "__main__":
    main(sys.argv[1:])