"""
Parser for DOCX footnotes and endnotes.
"""

import xml.etree.ElementTree as ET
from ..utils.xml_utils import qn, NSMAP


def parse_footnotes_xml(zipf):
    """
    Parses footnotes.xml to extract footnote content.
    
    Args:
        zipf: ZipFile object of the DOCX file
    
    Returns:
        dict: Mapping of footnote ID to footnote text
    """
    footnotes = {}
    
    try:
        footnotes_xml = zipf.read('word/footnotes.xml')
        root = ET.fromstring(footnotes_xml)
        
        for footnote in root.findall('.//{' + NSMAP['w'] + '}footnote'):
            footnote_id = footnote.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}id')
            
            # Extract text from paragraphs in footnote
            footnote_text = ''
            for para in footnote.findall('.//{' + NSMAP['w'] + '}p'):
                for run in para.findall('.//{' + NSMAP['w'] + '}r'):
                    for t in run.findall('.//{' + NSMAP['w'] + '}t'):
                        if t.text:
                            footnote_text += t.text
                    for br in run.findall('.//{' + NSMAP['w'] + '}br'):
                        footnote_text += '\n'
                footnote_text += '\n'
            
            if footnote_text.strip():
                footnotes[footnote_id] = footnote_text.strip()
    except (KeyError, ET.ParseError):
        # If footnotes.xml doesn't exist or can't be parsed
        pass
    
    return footnotes


def parse_endnotes_xml(zipf):
    """
    Parses endnotes.xml to extract endnote content.
    
    Args:
        zipf: ZipFile object of the DOCX file
    
    Returns:
        dict: Mapping of endnote ID to endnote text
    """
    endnotes = {}
    
    try:
        endnotes_xml = zipf.read('word/endnotes.xml')
        root = ET.fromstring(endnotes_xml)
        
        for endnote in root.findall('.//{' + NSMAP['w'] + '}endnote'):
            endnote_id = endnote.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}id')
            
            # Extract text from paragraphs in endnote
            endnote_text = ''
            for para in endnote.findall('.//{' + NSMAP['w'] + '}p'):
                for run in para.findall('.//{' + NSMAP['w'] + '}r'):
                    for t in run.findall('.//{' + NSMAP['w'] + '}t'):
                        if t.text:
                            endnote_text += t.text
                    for br in run.findall('.//{' + NSMAP['w'] + '}br'):
                        endnote_text += '\n'
                endnote_text += '\n'
            
            if endnote_text.strip():
                endnotes[endnote_id] = endnote_text.strip()
    except (KeyError, ET.ParseError):
        # If endnotes.xml doesn't exist or can't be parsed
        pass
    
    return endnotes
