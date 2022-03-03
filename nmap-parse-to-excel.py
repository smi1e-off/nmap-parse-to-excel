#!/usr/bin/python3
# -*- coding: utf-8 -*-
from doctest import script_from_examples
import xml.etree.ElementTree as ET
import sys
import os
import csv
import openpyxl
from libnmap.parser import NmapParser

    
def parseNmap(filename,out_filename):
    try:
        tree=ET.parse(filename)
        root=tree.getroot()
    except Exception as e:
        print (e)
    with open(out_filename,"w") as f:
        f.write("hostname, ip, port, status, protocol, service, product, version, extra info, OS type, vulners\n")
        for host in root.iter('host'):
            if host.find('status').get('state') == 'down':
                continue
            ip=host.find('address').get('addr',None)
            print("Reading information about %s！" % ip)
            try:
                if host.find('hostnames').find('hostname') is None:
                    hostname = ip
                else:
                    hostname = host.find('hostnames').find('hostname').get('name',None)
            except Exception as e:
                print("Error getting information about %s！" % ip)
                continue
            if not ip and not hostname:
                continue
            if host.find('ports').find('port') == None:
                output = hostname + "," + ip + ",,,,,,," + "\n"
                f.write(output)
            else:
                for ports in host.iter('port'):
                    port = ports.get('portid','')
                    status = ports.find('state').get('state','')
                    protocol = ports.get('protocol','')
   
                    if ports.find('service') != None:
                        service = ports.find('service').get('name','')
                        product = ports.find('service').get('product','')
                        version = ports.find('service').get('version','')
                        ostype = ports.find('service').get('ostype','')
                        extrainfo = ports.find('service').get('extrainfo','')
                    else:
                        service=''
                        product=''
                        version=''
                        ostype=''
                        extrainfo=''

                    for scripts in ports.findall('script'):
                        if scripts.findall('id') == 'vulners':
                            vulners = scripts.get('output')
                    else:
                        vulners=''
                        
                        output = hostname + "," + ip + "," + port + "," + status + "," +  protocol + "," + service + "," + product + "," + version + "," + extrainfo + "," + ostype + "," + vulners + "," "\n"
                        f.write(output)
    
def csv_to_excel(csv_file, excel_file):
    csv_data = []
    with open(csv_file) as file_obj:
        reader = csv.reader(file_obj)
        for row in reader:
            csv_data.append(row)
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    for row in csv_data:
        sheet.append(row)
    workbook.save(excel_file)
    
    
def main(args):
    for xml_file in args:
        print("deal with：" + xml_file)
        out_filename_csv = xml_file.rstrip(".xml") + ".csv"
        parseNmap(xml_file, out_filename_csv)
        out_filename_excel = out_filename_csv.rstrip(".csv") + ".xlsx"
        print(out_filename_excel)
        csv_to_excel(out_filename_csv, out_filename_excel)
        os.remove(out_filename_csv)
    
if __name__ == "__main__":
    if len(sys.argv[1:]) < 1:
        sys.exit("Instructions: %s 1.xml 2.xml 3.xml ...... " % __file__)
    
    else:
        main(sys.argv[1:])
    