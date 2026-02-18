"""
Parser for DOCX numbering.xml to understand list structures.
"""

import xml.etree.ElementTree as ET
from ..utils.xml_utils import NSMAP


def parse_numbering_xml(zipf):
    """
    Parses numbering.xml to understand list structures.
    
    Args:
        zipf: ZipFile object of the DOCX file
    
    Returns:
        dict: Mapping of numId to list information (list_type, num_format)
    """
    numbering_info = {}
    
    try:
        numbering_xml = zipf.read('word/numbering.xml')
        root = ET.fromstring(numbering_xml)
        
        # Find all num definitions
        for num in root.findall('.//{' + NSMAP['w'] + '}num'):
            numId = num.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}numId')
            abstractNumId_elem = num.find('{' + NSMAP['w'] + '}abstractNumId')
            
            if abstractNumId_elem is not None:
                abstractNumId = abstractNumId_elem.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val')
                
                # Find abstract numbering definition
                for abstractNum in root.findall('.//{' + NSMAP['w'] + '}abstractNum'):
                    if abstractNum.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}abstractNumId') == abstractNumId:
                        # Determine list type
                        list_type = 'bullet'
                        num_format = 'decimal'
                        
                        # Check for numbering format
                        for lvl in abstractNum.findall('.//{' + NSMAP['w'] + '}lvl'):
                            numFmt_elem = lvl.find('{' + NSMAP['w'] + '}numFmt')
                            if numFmt_elem is not None:
                                fmt_val = numFmt_elem.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', 'decimal')
                                if fmt_val == 'bullet':
                                    list_type = 'bullet'
                                else:
                                    list_type = 'number'
                                    num_format = fmt_val
                        
                        numbering_info[numId] = {
                            'list_type': list_type,
                            'num_format': num_format
                        }
                        break
    except (KeyError, ET.ParseError):
        # If numbering.xml doesn't exist or can't be parsed, use defaults
        pass
    
    return numbering_info
