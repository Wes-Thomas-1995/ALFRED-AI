import zipfile
from xml.etree import ElementTree as ET
import pandas as pd

# Constants
ROOT = r"Z:/38 - Private Debt/OPERATIONS/50_Projects/99_ESG/_EET/_EET_InternalProduction/2023 EET/_Production/ESG EET  I  3SC Pride Fund  I  2023 12.xlsm"
CUSTOM_XML_PATH = 'customXml/'
CONNECTIONS_XML_PATH = 'xl/connections.xml'
EXTRACT = []

def extract_power_query_code(zip_file):
    power_query_data = []
    
    # Step 1: Extract Power Query M code from customXml/item*.xml files
    for item in zip_file.namelist():
        if item.startswith(CUSTOM_XML_PATH) and item.endswith('.xml'):
            with zip_file.open(item) as custom_xml_file:
                try:
                    tree = ET.parse(custom_xml_file)
                    root = tree.getroot()
                    # Extract the content within the <xml> tag
                    power_query_code = ET.tostring(root, encoding='unicode')
                    power_query_data.append({
                        'File': item,
                        'PowerQueryCode': power_query_code.strip()
                    })
                except ET.ParseError:
                    print(f"Error parsing {item}")
    
    return power_query_data

def extract_query_metadata(zip_file):
    # Step 2: Extract metadata from xl/connections.xml
    metadata = []
    if CONNECTIONS_XML_PATH in zip_file.namelist():
        with zip_file.open(CONNECTIONS_XML_PATH) as connections_xml:
            tree = ET.parse(connections_xml)
            root = tree.getroot()
            ns = {'n': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
            
            for connection in root.findall('.//n:connection', ns):
                name = connection.get('name')
                conn_string = connection.find('n:connectionString', ns).text if connection.find('n:connectionString', ns) else ''
                metadata.append({'QueryName': name, 'ConnectionString': conn_string})
    
    return metadata

def extract_power_queries(root):
    with zipfile.ZipFile(root, 'r') as excel_zip:
        # Extract Power Query M code
        power_query_code_list = extract_power_query_code(excel_zip)
        
        # Extract metadata
        query_metadata_list = extract_query_metadata(excel_zip)
        
        # Combine into DataFrames
        power_query_df = pd.DataFrame(power_query_code_list)
        metadata_df = pd.DataFrame(query_metadata_list)
    
    return power_query_df, metadata_df

# Run the extraction
power_query_df, metadata_df = extract_power_queries(ROOT)

# Display results
print('\nPower Query Code:\n', power_query_df)
print('\nQuery Metadata:\n', metadata_df)