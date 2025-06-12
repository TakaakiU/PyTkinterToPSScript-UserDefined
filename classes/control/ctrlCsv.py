import os
import csv
import xml.etree.ElementTree as ET
from classes.control import ctrlMessage

class ctrlCsv():
    def read_csv(input_path):
        with open(input_path, newline='', encoding="utf-8-sig") as file:
            reader = csv.DictReader(file)
            rows = list(reader)
        return rows
        
    def output_csv(output_path, entry_data):
        _result = 0
        try:
            with open(output_path, mode="w", newline="", encoding="utf-8") as file:
                writer = csv.writer(file, quoting=csv.QUOTE_ALL)
                writer.writerows(entry_data)
        except Exception as err:
                _result = -3001
                ctrlMessage.print_error(err)
        
        return _result
    
    def extract_xml_to_csv(input_path, output_path):
        _result = 0
        tree = ET.parse(input_path)
        root = tree.getroot()
        
        namespace = {"dsig": "http://www.w3.org/2000/09/xmldsig#"}

        try:
              with open(output_path, mode="w", newline="", encoding="utf-8") as file:
                   writer = csv.writer(file, quoting=csv.QUOTE_ALL)
                   writer.writerow(["No.", "Index", "HashValue"])

                   for index, reference in enumerate(root.findall("dsig:Reference", namespace), start=1):
                        uri = reference.get("URI")
                        digest_value = reference.find("dsig:DigestValue", namespace).text
                        writer.writerow([index, uri, digest_value])

        except Exception as err:
            _result = -3101
            ctrlMessage.print_error(err)
    
        return _result
    
    def extract_xmls_to_csv(filepath_lists, output_path):
        _result = 0

        # 存在する場合は、削除
        if os.path.exists(output_path):
            os.remove(output_path)
        
        try:
            with open(output_path, mode="w", newline="", encoding="utf-8") as file:
                writer = csv.writer(file, quoting=csv.QUOTE_ALL)
                writer.writerow(["No.", "Index", "HashValue"])

            _total_count = 1

            for input_path in filepath_lists: 
                _xml_path = input_path + '/META-INF/Manifest.xml'
                tree = ET.parse(_xml_path)
                root = tree.getroot()
                
                namespace = {"dsig": "http://www.w3.org/2000/09/xmldsig#"}
                
                with open(output_path, mode="a", newline="", encoding="utf-8") as file:
                    writer = csv.writer(file, quoting=csv.QUOTE_ALL)
                    
                    for index, reference in enumerate(root.findall("dsig:Reference", namespace), start=1):
                        uri = reference.get("URI")
                        _root_path = os.path.basename(input_path)
                        uri = _root_path + '/' + uri
                        digest_value = reference.find("dsig:DigestValue", namespace).text
                        writer.writerow([_total_count, uri, digest_value])
                        _total_count += 1
        
        except Exception as err:
            _result = -3201
            ctrlMessage.print_error(err)
        
        return _result
