import re
import ast
import zipfile
import openpyxl
import pandas as pd
import networkx as nx
import xml.etree.ElementTree as ET
from collections import defaultdict






def DEPENDENCY_GRAPHING(FORMULAS):

    G = nx.DiGraph()

    for SHEET_NAME, CELLS in FORMULAS.items():
        for CELL, CELL_DATA in CELLS.items():
            NODE_ID = f"{SHEET_NAME}!{CELL}"
            G.add_node(NODE_ID)

            # Add edges for each reference in the formula
            REFERENCES = CELL_DATA.get('REFERENCES', [])
            for REF in REFERENCES:
                REF_SHEET = REF['SHEET']
                REF_CELL = REF['CELL']
                REF_NODE_ID = f"{REF_SHEET}!{REF_CELL}"
                G.add_node(REF_NODE_ID)  
                G.add_edge(REF_NODE_ID, NODE_ID)

    return G




def FORMULA_CLUSTERS(FORMULA):
    
    CLUSTERS = defaultdict(dict)

    for SHEET_NAME, CELLS in FORMULA.items():
        FORM_CLUSTERS = defaultdict(list)

        for CELL, CELL_DATA in CELLS.items():
            FORMULA_KEY = CELL_DATA.get('FORMULA', 'NO_FORMULA')
            FORM_CLUSTERS[FORMULA_KEY].append(CELL)

        CLUSTERS[SHEET_NAME] = dict(FORM_CLUSTERS)

    return CLUSTERS




def MAP_CROSS_REFERENCES(CROSS_SHEET_REF, WORKSHEETS):

    CROSS_SHEET_MAP = defaultdict(list)

    for SHEET_NAME, CELLS in CROSS_SHEET_REF.items():
        for CELL, REFERENCES in CELLS.items():
            for REF in REFERENCES:
                REF_SHEET   = REF['SHEET']
                REF_CELL    = REF['CELL']

                if REF_SHEET in WORKSHEETS:
                    CROSS_SHEET_MAP[(SHEET_NAME, CELL)].append((REF_SHEET, REF_CELL))

    return CROSS_SHEET_MAP







def PHASE_3_OVERVIEW(FORMULA_REFS, CROSS_SHEET_REFS):

    DEPENDENCY_GRAPH    = DEPENDENCY_GRAPHING(FORMULA_REFS)
    FORMULA_CLUSTER     = FORMULA_CLUSTERS(FORMULA_REFS)
    CROSS_SHEET_MAP     = MAP_CROSS_REFERENCES(CROSS_SHEET_REFS, FORMULA_REFS)

    return {'DEPENDENCY_GRAPH'  : DEPENDENCY_GRAPH,
            'CLUSTERS'          : FORMULA_CLUSTER,
            'CROSS_SHEET_MAP'   : CROSS_SHEET_MAP}





