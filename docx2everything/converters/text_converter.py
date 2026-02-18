"""
Plain text converter for DOCX files.
"""

import xml.etree.ElementTree as ET
from ..utils.xml_utils import qn


def xml2text(xml):
    """
    Converts XML content to plain text.
    
    A string representing the textual content of XML, with content
    child elements like ``<w:tab/>`` translated to their Python
    equivalent.
    
    Args:
        xml: XML content as bytes or string
    
    Returns:
        Plain text string
    """
    text = ''
    root = ET.fromstring(xml)
    for child in root.iter():
        if child.tag == qn('w:t'):
            t_text = child.text
            text += t_text if t_text is not None else ''
        elif child.tag == qn('w:tab'):
            text += '\t'
        elif child.tag in (qn('w:br'), qn('w:cr')):
            text += '\n'
        elif child.tag == qn("w:p"):
            text += '\n\n'
    return text


def convert_to_text(zipf, filelist, img_dir=None):
    """
    Converts DOCX file to plain text.
    
    Args:
        zipf: ZipFile object of the DOCX file
        filelist: List of files in the ZIP archive
        img_dir: Optional directory to extract images
    
    Returns:
        Plain text string
    """
    text = ''
    
    # Get header text
    header_xmls = 'word/header[0-9]*.xml'
    import re
    for fname in filelist:
        if re.match(header_xmls, fname):
            text += xml2text(zipf.read(fname))
    
    # Get main text
    doc_xml = 'word/document.xml'
    text += xml2text(zipf.read(doc_xml))
    
    # Get footer text
    footer_xmls = 'word/footer[0-9]*.xml'
    for fname in filelist:
        if re.match(footer_xmls, fname):
            text += xml2text(zipf.read(fname))
    
    # Extract images if needed
    if img_dir is not None:
        from ..utils.file_utils import extract_images
        extract_images(zipf, filelist, img_dir)
    
    return text.strip()
