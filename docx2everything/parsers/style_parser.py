"""
Parser for DOCX styles.xml to understand custom styles and formatting.
"""

import xml.etree.ElementTree as ET
from ..utils.xml_utils import NSMAP


def parse_styles_xml(zipf):
    """
    Parses styles.xml to extract style information.
    
    Args:
        zipf: ZipFile object of the DOCX file
    
    Returns:
        dict: Mapping of style ID to style information (name, type, based_on, etc.)
    """
    styles = {}
    
    try:
        styles_xml = zipf.read('word/styles.xml')
        root = ET.fromstring(styles_xml)
        
        for style in root.findall('.//{' + NSMAP['w'] + '}style'):
            style_id = style.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}styleId')
            style_type = style.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}type')
            
            style_info = {
                'type': style_type,
                'name': None,
                'based_on': None,
                'is_heading': False,
                'heading_level': None
            }
            
            # Get style name
            name_elem = style.find('{' + NSMAP['w'] + '}name')
            if name_elem is not None:
                style_info['name'] = name_elem.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val')
            
            # Get basedOn style
            based_on_elem = style.find('{' + NSMAP['w'] + '}basedOn')
            if based_on_elem is not None:
                style_info['based_on'] = based_on_elem.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val')
            
            # Check if it's a heading style
            if style_info['name']:
                name_lower = style_info['name'].lower()
                if 'heading' in name_lower or 'title' in name_lower:
                    style_info['is_heading'] = True
                    # Try to extract heading level
                    for i in range(1, 7):
                        if f'heading{i}' in name_lower or f'heading {i}' in name_lower:
                            style_info['heading_level'] = i
                            break
                    if 'title' in name_lower and style_info['heading_level'] is None:
                        style_info['heading_level'] = 1
            
            # Also check based_on style recursively
            if style_info['based_on'] and style_info['based_on'] in styles:
                based_on_info = styles[style_info['based_on']]
                if based_on_info.get('is_heading'):
                    style_info['is_heading'] = True
                    style_info['heading_level'] = based_on_info.get('heading_level')
            
            if style_id:
                styles[style_id] = style_info
    except (KeyError, ET.ParseError):
        # If styles.xml doesn't exist or can't be parsed
        pass
    
    return styles
