import zipfile
import xml.etree.ElementTree as ET
from collections import defaultdict





class EXCEL_XML_EXTRACTOR:
    def __init__(self, FILE_PATH):
        self.FILE_PATH = FILE_PATH
        self.XML_FILES = self.EXTRACT_XML_FILES()

    def EXTRACT_XML_FILES(self):
        XML_FILES = {}
        with zipfile.ZipFile(self.FILE_PATH, 'r') as ZIP_REF:
            for FILE in ZIP_REF.namelist():
                if FILE.endswith('.xml'):
                    with ZIP_REF.open(FILE) as f:
                        XML_FILES[FILE] = f.read()
        return XML_FILES







class WORKBOOK_PARSING:
    def __init__(self, XML_FILES):
        self.XML_FILES = XML_FILES
        self.DICT = self.PARSE_WORKBOOK()

    def PARSE_WORKBOOK(self):
        MAP = {}
        if 'xl/workbook.xml' in self.XML_FILES:
            ROOT = ET.fromstring(self.XML_FILES['xl/workbook.xml'])
            SHEETS = ROOT.findall('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}sheet')
            for i, SHEET in enumerate(SHEETS):
                NAME = SHEET.attrib.get('name')
                STATE = SHEET.attrib.get('state', 'visible')
                MAP[f'sheet{i + 1}'] = {'NAME': NAME, 'STATE': STATE}
        return MAP









class WORKSHEET_PARSING:
    def __init__(self, XML_FILES, SHEET_NAME_MAP):
        self.XML_FILES = XML_FILES
        self.SHEET_NAME_MAP = SHEET_NAME_MAP
        self.DICT = self.PARSE_WORKSHEETS()

    def PARSE_WORKSHEETS(self):
        WORKSHEETS = {}
        for FILE_NAME, XML_CONTENT in self.XML_FILES.items():
            if 'xl/worksheets/sheet' in FILE_NAME:
                ROOT = ET.fromstring(XML_CONTENT)
                SHEET_DATA = {}

                for CELL in ROOT.findall('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}c'):
                    CELL_REF = CELL.attrib.get('r')
                    CELL_TYPE = CELL.attrib.get('t')
                    STYLE_INDEX = CELL.attrib.get('s')
                    FORMULA = CELL.find('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}f')
                    VALUE = CELL.find('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}v')

                    VALUE_TEXT = VALUE.text if VALUE is not None else None

                    CELL_DATA = {
                        'TYPE': CELL_TYPE,
                        'STYLE_INDEX': STYLE_INDEX,
                        'FORMULA': FORMULA.text if FORMULA is not None else None,
                        'VALUE': VALUE_TEXT
                    }

                    SHEET_DATA[CELL_REF] = CELL_DATA

                # Extract the sheet internal reference (e.g., 'sheet1')
                SHEET_NAME_KEY = FILE_NAME.split('/')[-1].replace('.xml', '')
                # Map the internal sheet name to the user-defined sheet name
                USER_DEFINED_NAME = self.SHEET_NAME_MAP.get(SHEET_NAME_KEY, {}).get('NAME', SHEET_NAME_KEY)
                WORKSHEETS[USER_DEFINED_NAME] = SHEET_DATA

        return WORKSHEETS





class CALC_CHAIN_PARSING:
    def __init__(self, XML_FILES, SHEET_NAME_MAP):
        self.XML_FILES = XML_FILES
        self.SHEET_NAME_MAP = SHEET_NAME_MAP
        self.DICT = self.CALC_CHAIN_PARSE()

    def CALC_CHAIN_PARSE(self):
        CALC_CHAIN = []
        if 'xl/calcChain.xml' in self.XML_FILES:
            ROOT = ET.fromstring(self.XML_FILES['xl/calcChain.xml'])
            for CELL in ROOT.findall('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}c'):
                CELL_REF = CELL.attrib.get('r')
                SHEET_ID = CELL.attrib.get('i')
                # Map SHEET_ID to the user-defined sheet name using SHEET_NAME_MAP
                SHEET_NAME = self.SHEET_NAME_MAP.get(f'sheet{SHEET_ID}', {}).get('NAME', f'sheet{SHEET_ID}')
                CALC_CHAIN.append({'CELL_REF': CELL_REF, 'SHEET_NAME': SHEET_NAME})
        return CALC_CHAIN





class SHARED_STRING_PARSING:

    def __init__(self, XML_FILES):
        self.XML_FILES = XML_FILES
        self.DICT = self.SHARED_STRINGS_PARSE()

    def SHARED_STRINGS_PARSE(self):
        SHARED_STRINGS = {}
        if 'xl/sharedStrings.xml' in self.XML_FILES:
            ROOT = ET.fromstring(self.XML_FILES['xl/sharedStrings.xml'])
            for i, SI in enumerate(ROOT.findall('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}si')):
                TEXT_NODE = SI.find('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t')
                if TEXT_NODE is not None:
                    SHARED_STRINGS[i] = TEXT_NODE.text
        return SHARED_STRINGS








class STYLES_PARSING:
    
    def __init__(self, XML_FILES):
        self.XML_FILES = XML_FILES
        self.DICT = self.STYLE_PARSE()

    def STYLE_PARSE(self):
        STYLES = {  
            'NUMBER_FORMATS': {},
            'FONTS': [],
            'FILLS': [],
            'BORDERS': [],
            'CELLXFS': []
        }
        
        if 'xl/styles.xml' in self.XML_FILES:
            ROOT = ET.fromstring(self.XML_FILES['xl/styles.xml'])

            # 1. Extract number formats
            NUM_FMTS = ROOT.find('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}numFmts')
            if NUM_FMTS is not None:
                for NUM_FMT in NUM_FMTS.findall('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}numFmt'):
                    FMT_ID = NUM_FMT.attrib.get('numFmtId')
                    FORMAT_CODE = NUM_FMT.attrib.get('formatCode')
                    STYLES['NUMBER_FORMATS'][FMT_ID] = FORMAT_CODE

            # 2. Extract fonts
            FONTS = ROOT.find('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}fonts')
            if FONTS is not None:
                for FONT in FONTS.findall('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}font'):
                    FONT_DATA = {
                        'NAME': FONT.find('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}name').attrib.get('val'),
                        'SIZE': FONT.find('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}sz').attrib.get('val'),
                        'BOLD': FONT.find('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}b') is not None,
                        'ITALIC': FONT.find('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}i') is not None,
                        'COLOR': FONT.find('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}color').attrib.get('rgb') if FONT.find('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}color') is not None else None
                    }
                    STYLES['FONTS'].append(FONT_DATA)

            # 3. Extract fills
            FILLS = ROOT.find('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}fills')
            if FILLS is not None:
                for FILL in FILLS.findall('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}fill'):
                    PATTERN_FILL = FILL.find('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}patternFill')
                    FILL_DATA = {
                        'PATTERN_TYPE': PATTERN_FILL.attrib.get('patternType') if PATTERN_FILL is not None else None,
                        'FG_COLOR': PATTERN_FILL.find('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}fgColor').attrib.get('rgb') if PATTERN_FILL is not None and PATTERN_FILL.find('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}fgColor') is not None else None,
                        'BG_COLOR': PATTERN_FILL.find('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}bgColor').attrib.get('rgb') if PATTERN_FILL is not None and PATTERN_FILL.find('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}bgColor') is not None else None
                    }
                    STYLES['FILLS'].append(FILL_DATA)

            # 4. Extract borders
            BORDERS = ROOT.find('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}borders')
            if BORDERS is not None:
                for BORDER in BORDERS.findall('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}border'):
                    BORDER_DATA = {
                        'LEFT': BORDER.find('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}left') is not None,
                        'RIGHT': BORDER.find('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}right') is not None,
                        'TOP': BORDER.find('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}top') is not None,
                        'BOTTOM': BORDER.find('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}bottom') is not None
                    }
                    STYLES['BORDERS'].append(BORDER_DATA)

            # 5. Extract cellXfs (Cell formats)
            CELLXFS = ROOT.find('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}cellXfs')
            if CELLXFS is not None:
                for XF in CELLXFS.findall('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}xf'):
                    XF_DATA = {
                        'NUM_FMT_ID': XF.attrib.get('numFmtId'),
                        'FONT_ID': XF.attrib.get('fontId'),
                        'FILL_ID': XF.attrib.get('fillId'),
                        'BORDER_ID': XF.attrib.get('borderId')
                    }
                    STYLES['CELLXFS'].append(XF_DATA)

        return STYLES









class DRAWING_PARSING:

    def __init__(self, XML_FILES):
        self.XML_FILES = XML_FILES
        self.DICT = self.DRAWING_PARSE()

    def DRAWING_PARSE(self):
        DRAWINGS = defaultdict(list)
        for FILE_NAME, XML_CONTENT in self.XML_FILES.items():
            if 'xl/drawings/drawing' in FILE_NAME:
                ROOT = ET.fromstring(XML_CONTENT)
                DRAWING_INFO = []

                # Look for twoCellAnchor elements
                for ANCHOR in ROOT.findall('.//{http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing}twoCellAnchor'):
                    
                    # Identify images
                    BLIP = ANCHOR.find('.//{http://schemas.openxmlformats.org/drawingml/2006/main}blip')
                    if BLIP is not None:
                        EMBED = BLIP.attrib.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed')
                        DRAWING_INFO.append({'TYPE': 'IMAGE', 'EMBED_ID': EMBED})

                    # Identify shapes (e.g., rectangles, circles)
                    SP = ANCHOR.find('.//{http://schemas.openxmlformats.org/drawingml/2006/main}sp')
                    if SP is not None:
                        SP_NAME = SP.find('.//{http://schemas.openxmlformats.org/drawingml/2006/main}nvSpPr/{http://schemas.openxmlformats.org/drawingml/2006/main}cNvPr').attrib.get('name', 'Shape')
                        DRAWING_INFO.append({'TYPE': 'SHAPE', 'NAME': SP_NAME})

                    # Identify charts
                    GRAPHIC_FRAME = ANCHOR.find('.//{http://schemas.openxmlformats.org/drawingml/2006/main}graphicFrame')
                    if GRAPHIC_FRAME is not None:
                        CHART = GRAPHIC_FRAME.find('.//{http://schemas.openxmlformats.org/drawingml/2006/chart}chart')
                        if CHART is not None:
                            CHART_ID = CHART.attrib.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
                            DRAWING_INFO.append({'TYPE': 'CHART', 'CHART_ID': CHART_ID})

                    # Identify connectors (e.g., lines)
                    CNX_SP = ANCHOR.find('.//{http://schemas.openxmlformats.org/drawingml/2006/main}cxnSp')
                    if CNX_SP is not None:
                        CNX_NAME = CNX_SP.find('.//{http://schemas.openxmlformats.org/drawingml/2006/main}nvCxnSpPr/{http://schemas.openxmlformats.org/drawingml/2006/main}cNvPr').attrib.get('name', 'Connector')
                        DRAWING_INFO.append({'TYPE': 'CONNECTOR', 'NAME': CNX_NAME})

                DRAWINGS[FILE_NAME] = DRAWING_INFO
        return DRAWINGS



class TABLES_PARSING:

    def __init__(self, XML_FILES):
        self.XML_FILES = XML_FILES
        self.DICT = self.TABLES_PARSE()

    def TABLES_PARSE(self):

        TABLES = {}
        for FILE_NAME, XML_CONTENT in self.XML_FILES.items():
            if 'xl/tables/table' in FILE_NAME:
                ROOT = ET.fromstring(XML_CONTENT)
                TABLE_INFO = {}
                TABLE_INFO['NAME']      = ROOT.attrib.get('name')
                TABLE_INFO['REF']       = ROOT.attrib.get('ref')  # Table range (e.g., A1:B10)
                TABLE_INFO['COLUMNS']   = [col.attrib.get('name') for col in ROOT.findall('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}tableColumn')]
                TABLES[FILE_NAME]       = TABLE_INFO
        return TABLES
    


class THEME_PARSING:

    def __init__(self, XML_FILES):
        self.XML_FILES = XML_FILES
        self.DICT = self.THEME_PARSE()

    def THEME_PARSE(self):

        THEME = {}
        if 'xl/theme/theme1.xml' in self.XML_FILES:  
            ROOT = ET.fromstring(self.XML_FILES['xl/theme/theme1.xml'])
            COLOUR_SCHEME = []
            for COLOUR in ROOT.findall('.//{http://schemas.openxmlformats.org/drawingml/2006/main}clrScheme//{http://schemas.openxmlformats.org/drawingml/2006/main}srgbClr'):
                COLOUR_SCHEME.append(COLOUR.attrib.get('val'))
            THEME['color_scheme'] = COLOUR_SCHEME
        return THEME




class RELATIVE_SHEETS_PARSING:
    
    def __init__(self, XML_FILES, SHEET_NAME_MAP):
        self.XML_FILES = XML_FILES
        self.SHEET_NAME_MAP = SHEET_NAME_MAP
        self.DICT = self.parse_sheet_rels()

    def parse_sheet_rels(self):
        SHEET_RELS = defaultdict(list)
        
        for FILE_NAME, XML_CONTENT in self.XML_FILES.items():
            if 'xl/worksheets/_rels/sheet' in FILE_NAME and '.rels' in FILE_NAME:
                ROOT = ET.fromstring(XML_CONTENT)

                SHEET_NUM_KEY = FILE_NAME.split('/')[-1].replace('.xml.rels', '')
                USER_DEFINED_NAME = self.SHEET_NAME_MAP.get(SHEET_NUM_KEY, {}).get('NAME', SHEET_NUM_KEY)

                for REL in ROOT.findall('.//{http://schemas.openxmlformats.org/package/2006/relationships}Relationship'):
                    REL_ID = REL.attrib.get('Id')
                    TARGET = REL.attrib.get('Target')
                    REL_TYPE = REL.attrib.get('Type').split('/')[-1].upper()  # Capitalize the TYPE
                    SHEET_RELS[USER_DEFINED_NAME].append({'ID': REL_ID, 'TARGET': TARGET, 'TYPE': REL_TYPE})

        return SHEET_RELS
    



class CONTENT_TYPE_PARSING:

    def __init__(self, XML_FILES):
        self.XML_FILES = XML_FILES
        self.DICT = self.CONTENT_TYPE_PARSE()

    def CONTENT_TYPE_PARSE(self):

        CONTENT_TYPES = []
        if '[Content_Types].xml' in self.XML_FILES:
            ROOT = ET.fromstring(self.XML_FILES['[Content_Types].xml'])
            for OVERRIDE in ROOT.findall('.//{http://schemas.openxmlformats.org/package/2006/content-types}Override'):
                PART_NAME       = OVERRIDE.attrib.get('PartName')
                CONTENT_TYPE    = OVERRIDE.attrib.get('ContentType')
                CONTENT_TYPES.append({'PART_NAME': PART_NAME, 'CONTENT_TYPE': CONTENT_TYPE})
        return CONTENT_TYPES





class RELATIVE_WORKBOOK_PARSING:
    
    def __init__(self, XML_FILES):
        self.XML_FILES = XML_FILES
        self.DICT = self.WORKBOOK_REL_PARSE()

    def WORKBOOK_REL_PARSE(self):

        WORKBOOK_RELS = []
        if 'xl/_rels/workbook.xml.rels' in self.XML_FILES:
            ROOT = ET.fromstring(self.XML_FILES['xl/_rels/workbook.xml.rels'])
            for REL in ROOT.findall('.//{http://schemas.openxmlformats.org/package/2006/relationships}Relationship'):
                REL_ID      = REL.attrib.get('Id')
                TARGET      = REL.attrib.get('Target')
                REL_TYPE    = REL.attrib.get('Type').split('/')[-1]
                WORKBOOK_RELS.append({'ID': REL_ID, 'TARGET': TARGET, 'TYPE': REL_TYPE})
        return WORKBOOK_RELS







class EXCEL_DATA_PARSER:


    def __init__(self, FILE_PATH):
        self.XML_EXTRACTOR      = EXCEL_XML_EXTRACTOR(FILE_PATH)
        self.SHEET_NAMES        = WORKBOOK_PARSING(self.XML_EXTRACTOR.XML_FILES)
        self.WORKSHEET_PARSER   = WORKSHEET_PARSING(self.XML_EXTRACTOR.XML_FILES, self.SHEET_NAMES.DICT)
        self.CALC_CHAIN         = CALC_CHAIN_PARSING(self.XML_EXTRACTOR.XML_FILES, self.SHEET_NAMES.DICT)
        self.SHEET_RELS         = RELATIVE_SHEETS_PARSING(self.XML_EXTRACTOR.XML_FILES, self.SHEET_NAMES.DICT)
        self.SHARED_STRINGS     = SHARED_STRING_PARSING(self.XML_EXTRACTOR.XML_FILES)
        self.STYLES             = STYLES_PARSING(self.XML_EXTRACTOR.XML_FILES)
        self.DRAWINGS           = DRAWING_PARSING(self.XML_EXTRACTOR.XML_FILES)
        self.TABLES             = TABLES_PARSING(self.XML_EXTRACTOR.XML_FILES)
        self.THEME              = THEME_PARSING(self.XML_EXTRACTOR.XML_FILES)
        self.CONTENT_TYPE       = CONTENT_TYPE_PARSING(self.XML_EXTRACTOR.XML_FILES)
        self.WORKBOOK_RELS      = RELATIVE_WORKBOOK_PARSING(self.XML_EXTRACTOR.XML_FILES)


    def GET_DATA(self):

        return {'WORKSHEETS'        : self.WORKSHEET_PARSER.DICT,
                'SHEET_NAME_MAP'    : self.SHEET_NAMES.DICT,
                'SHARED_STRINGS'    : self.SHARED_STRINGS.DICT,
                'STYLES'            : self.STYLES.DICT,
                'CALC_CHAIN'        : self.CALC_CHAIN.DICT,
                'DRAWINGS'          : self.DRAWINGS.DICT,
                'TABLES'            : self.TABLES.DICT,
                'SHEET_RELS'        : self.SHEET_RELS.DICT,
                'THEME'             : self.THEME.DICT,
                'CONTENT_TYPES'     : self.CONTENT_TYPE.DICT,
                'WORKBOOK_RELS'     : self.WORKBOOK_RELS.DICT}





