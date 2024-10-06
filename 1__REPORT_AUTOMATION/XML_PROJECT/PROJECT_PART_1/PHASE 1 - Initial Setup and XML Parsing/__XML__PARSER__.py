import zipfile
import xml.etree.ElementTree as ET
from collections import defaultdict

class WorkbookParser:
    """
    A class to parse the workbook.xml to extract sheet names and properties.
    """

    def __init__(self, xml_files):
        self.xml_files = xml_files
        self.sheets_info = self.parse_workbook()

    def parse_workbook(self):
        """
        Parses xl/workbook.xml to extract sheet names and properties.
        Returns a dictionary of sheet names with their attributes (e.g., ID, visibility).
        """
        sheets_info = {}
        if 'xl/workbook.xml' in self.xml_files:
            root = ET.fromstring(self.xml_files['xl/workbook.xml'])
            for sheet in root.findall('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}sheet'):
                sheet_id = sheet.attrib.get('sheetId')
                name = sheet.attrib.get('name')
                state = sheet.attrib.get('state', 'visible')  # Default to 'visible' if not present
                sheets_info[sheet_id] = {'name': name, 'state': state}
        return sheets_info


class CalcChainParser:
    """
    A class to parse the calcChain.xml to extract the calculation order.
    """

    def __init__(self, xml_files):
        self.xml_files = xml_files
        self.calc_chain = self.parse_calc_chain()

    def parse_calc_chain(self):
        """
        Parses xl/calcChain.xml to extract the calculation sequence of cells.
        Returns a list of cell references in the order they are calculated.
        """
        calc_chain = []
        if 'xl/calcChain.xml' in self.xml_files:
            root = ET.fromstring(self.xml_files['xl/calcChain.xml'])
            for cell in root.findall('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}c'):
                cell_ref = cell.attrib.get('r')
                sheet_id = cell.attrib.get('i')  # Sheet ID
                calc_chain.append({'cell_ref': cell_ref, 'sheet_id': sheet_id})
        return calc_chain


class DrawingsParser:
    """
    A class to parse drawing XML files to extract drawing information.
    """

    def __init__(self, xml_files):
        self.xml_files = xml_files
        self.drawings = self.parse_drawings()

    def parse_drawings(self):
        """
        Parses xl/drawings/drawingX.xml to extract shapes, images, and charts.
        Returns a list of drawing elements (e.g., images, shapes) with their properties.
        """
        drawings = defaultdict(list)
        for file_name, xml_content in self.xml_files.items():
            if 'xl/drawings/drawing' in file_name:
                root = ET.fromstring(xml_content)
                drawing_info = []
                for anchor in root.findall('.//{http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing}twoCellAnchor'):
                    blip = anchor.find('.//{http://schemas.openxmlformats.org/drawingml/2006/main}blip')
                    if blip is not None:
                        embed = blip.attrib.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed')
                        drawing_info.append({'type': 'image', 'embed_id': embed})
                drawings[file_name] = drawing_info
        return drawings


class TablesParser:
    """
    A class to parse table XML files to extract table structures and metadata.
    """

    def __init__(self, xml_files):
        self.xml_files = xml_files
        self.tables = self.parse_tables()

    def parse_tables(self):
        """
        Parses xl/tables/tableX.xml to extract table metadata.
        Returns a dictionary of tables with their columns and ranges.
        """
        tables = {}
        for file_name, xml_content in self.xml_files.items():
            if 'xl/tables/table' in file_name:
                root = ET.fromstring(xml_content)
                table_info = {}
                table_info['name'] = root.attrib.get('name')
                table_info['ref'] = root.attrib.get('ref')  # Table range (e.g., A1:B10)
                table_info['columns'] = [col.attrib.get('name') for col in root.findall('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}tableColumn')]
                tables[file_name] = table_info
        return tables


class SheetRelsParser:
    """
    A class to parse sheet relationships (sheetX.xml.rels) to identify linked objects.
    """

    def __init__(self, xml_files):
        self.xml_files = xml_files
        self.sheet_rels = self.parse_sheet_rels()

    def parse_sheet_rels(self):
        """
        Parses xl/worksheets/_rels/sheetX.xml.rels to extract relationships (e.g., images, charts).
        Returns a dictionary of relationships for each sheet.
        """
        sheet_rels = defaultdict(list)
        for file_name, xml_content in self.xml_files.items():
            if 'xl/worksheets/_rels/sheet' in file_name and '.rels' in file_name:
                root = ET.fromstring(xml_content)
                for rel in root.findall('.//{http://schemas.openxmlformats.org/package/2006/relationships}Relationship'):
                    rel_id = rel.attrib.get('Id')
                    target = rel.attrib.get('Target')
                    rel_type = rel.attrib.get('Type').split('/')[-1]  # Extract relationship type
                    sheet_rels[file_name].append({'id': rel_id, 'target': target, 'type': rel_type})
        return sheet_rels


class ThemeParser:
    """
    A class to parse theme XML files to extract color schemes and fonts.
    """

    def __init__(self, xml_files):
        self.xml_files = xml_files
        self.theme = self.parse_theme()

    def parse_theme(self):
        """
        Parses xl/theme/themeX.xml to extract color schemes and fonts.
        Returns a dictionary containing theme details.
        """
        theme = {}
        if 'xl/theme/theme1.xml' in self.xml_files:  # Default theme is usually theme1.xml
            root = ET.fromstring(self.xml_files['xl/theme/theme1.xml'])
            color_scheme = []
            for color in root.findall('.//{http://schemas.openxmlformats.org/drawingml/2006/main}clrScheme//{http://schemas.openxmlformats.org/drawingml/2006/main}srgbClr'):
                color_scheme.append(color.attrib.get('val'))
            theme['color_scheme'] = color_scheme
        return theme


class ContentTypesParser:
    """
    A class to parse [Content_Types].xml to identify parts of the document and their types.
    """

    def __init__(self, xml_files):
        self.xml_files = xml_files
        self.content_types = self.parse_content_types()

    def parse_content_types(self):
        """
        Parses [Content_Types].xml to identify content types within the workbook.
        Returns a list of content types and their paths.
        """
        content_types = []
        if '[Content_Types].xml' in self.xml_files:
            root = ET.fromstring(self.xml_files['[Content_Types].xml'])
            for override in root.findall('.//{http://schemas.openxmlformats.org/package/2006/content-types}Override'):
                part_name = override.attrib.get('PartName')
                content_type = override.attrib.get('ContentType')
                content_types.append({'part_name': part_name, 'content_type': content_type})
        return content_types


class WorkbookRelsParser:
    """
    A class to parse workbook.xml.rels to identify relationships in the workbook.
    """

    def __init__(self, xml_files):
        self.xml_files = xml_files
        self.workbook_rels = self.parse_workbook_rels()

    def parse_workbook_rels(self):
        """
        Parses xl/_rels/workbook.xml.rels to extract workbook relationships.
        Returns a list of relationships with their targets and types.
        """
        workbook_rels = []
        if 'xl/_rels/workbook.xml.rels' in self.xml_files:
            root = ET.fromstring(self.xml_files['xl/_rels/workbook.xml.rels'])
            for rel in root.findall('.//{http://schemas.openxmlformats.org/package/2006/relationships}Relationship'):
                rel_id = rel.attrib.get('Id')
                target = rel.attrib.get('Target')
                rel_type = rel.attrib.get('Type').split('/')[-1]
                workbook_rels.append({'id': rel_id, 'target': target, 'type': rel_type})
        return workbook_rels



class ExcelXMLExtractor:
    """
    A class to extract XML files from an Excel workbook.
    """

    def __init__(self, file_path):
        self.file_path = file_path
        self.xml_files = self.extract_xml_files()

    def extract_xml_files(self):
        """
        Extracts XML files from the Excel workbook.
        Returns a dictionary with filenames as keys and XML content as values.
        """
        xml_files = {}
        with zipfile.ZipFile(self.file_path, 'r') as zip_ref:
            for file in zip_ref.namelist():
                if file.endswith('.xml'):
                    with zip_ref.open(file) as f:
                        xml_files[file] = f.read()
        return xml_files


class WorksheetParser:
    """
    A class to parse worksheet XML files to extract cell data and formulas.
    """

    def __init__(self, xml_files):
        self.xml_files = xml_files
        self.worksheets = self.parse_worksheets()

    def parse_worksheets(self):
        """
        Parses worksheet XML files to extract cell data and formulas.
        Returns a dictionary with worksheet names as keys and cell data as values.
        """
        worksheets = {}
        for file_name, xml_content in self.xml_files.items():
            if 'xl/worksheets/sheet' in file_name:
                root = ET.fromstring(xml_content)
                sheet_data = {}

                for cell in root.findall('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}c'):
                    cell_ref = cell.attrib.get('r')
                    cell_type = cell.attrib.get('t')
                    formula = cell.find('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}f')
                    value = cell.find('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}v')

                    cell_data = {
                        'type': cell_type,
                        'formula': formula.text if formula is not None else None,
                        'value': value.text if value is not None else None
                    }
                    sheet_data[cell_ref] = cell_data

                sheet_name = file_name.split('/')[-1].replace('.xml', '')
                worksheets[sheet_name] = sheet_data

        return worksheets


class SharedStringsParser:
    """
    A class to parse shared strings XML to extract text values.
    """

    def __init__(self, xml_files):
        self.xml_files = xml_files
        self.shared_strings = self.parse_shared_strings()

    def parse_shared_strings(self):
        """
        Parses shared strings XML to extract text values.
        Returns a dictionary with string indices as keys and text content as values.
        """
        shared_strings = {}
        if 'xl/sharedStrings.xml' in self.xml_files:
            root = ET.fromstring(self.xml_files['xl/sharedStrings.xml'])
            for i, si in enumerate(root.findall('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}si')):
                text_node = si.find('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t')
                if text_node is not None:
                    shared_strings[i] = text_node.text
        return shared_strings


class StylesParser:
    """
    A placeholder class to parse styles from the styles XML file.
    """

    def __init__(self, xml_files):
        self.xml_files = xml_files
        self.styles = self.parse_styles()

    def parse_styles(self):
        """
        Placeholder method for parsing styles from styles.xml.
        Returns a dictionary mapping style IDs to style attributes.
        """
        styles = {}
        if 'xl/styles.xml' in self.xml_files:
            # Placeholder for extracting styles like number formats, fonts, colors
            pass
        return styles



# Controller Class
class ExcelDataParser:
    def __init__(self, file_path):
        self.xml_extractor = ExcelXMLExtractor(file_path)
        self.xml_files = self.xml_extractor.xml_files

        # Initializing all parsers
        self.worksheet_parser = WorksheetParser(self.xml_files)
        self.shared_strings_parser = SharedStringsParser(self.xml_files)
        self.styles_parser = StylesParser(self.xml_files)
        self.workbook_parser = WorkbookParser(self.xml_files)
        self.calc_chain_parser = CalcChainParser(self.xml_files)
        self.drawings_parser = DrawingsParser(self.xml_files)
        self.tables_parser = TablesParser(self.xml_files)
        self.sheet_rels_parser = SheetRelsParser(self.xml_files)
        self.theme_parser = ThemeParser(self.xml_files)
        self.content_types_parser = ContentTypesParser(self.xml_files)
        self.workbook_rels_parser = WorkbookRelsParser(self.xml_files)

    def get_all_data(self):
        """
        Orchestrates all parsing processes and returns a dictionary with all parsed data.
        """
        return {
            'worksheets': self.worksheet_parser.worksheets,
            'shared_strings': self.shared_strings_parser.shared_strings,
            'styles': self.styles_parser.styles,
            'sheets_info': self.workbook_parser.sheets_info,
            'calc_chain': self.calc_chain_parser.calc_chain,
            'drawings': self.drawings_parser.drawings,
            'tables': self.tables_parser.tables,
            'sheet_rels': self.sheet_rels_parser.sheet_rels,
            'theme': self.theme_parser.theme,
            'content_types': self.content_types_parser.content_types,
            'workbook_rels': self.workbook_rels_parser.workbook_rels
        }




# Usage example
# This part initializes the ExcelDataParser with an Excel file path and extracts all relevant data.
# excel_data_parser = ExcelDataParser('path/to/excel.xlsx')
# all_data = excel_data_parser.get_all_data()
# print(all_data)

# Example usage (replace 'path/to/excel.xlsx' with the actual path to your Excel file)
# excel_data_parser = ExcelDataParser('path/to/excel.xlsx')
# parsed_data = excel_data_parser.get_data()
# print(parsed_data)