"""
Parser for DOCX relationship files to resolve hyperlinks and images.
"""

import xml.etree.ElementTree as ET


def parse_relationships(zipf, rel_file='word/_rels/document.xml.rels'):
    """
    Parses relationship files to resolve hyperlinks and images.
    
    Args:
        zipf: ZipFile object of the DOCX file
        rel_file: Path to the relationship file within the DOCX
    
    Returns:
        tuple: (hyperlinks_dict, images_dict) mapping relationship IDs to URLs/paths
    """
    hyperlinks = {}
    images = {}
    
    try:
        rels_xml = zipf.read(rel_file)
        root = ET.fromstring(rels_xml)
        
        # Relationship elements are in the package namespace
        rel_ns = 'http://schemas.openxmlformats.org/package/2006/relationships'
        
        for rel in root.findall('.//{' + rel_ns + '}Relationship'):
            rel_id = rel.get('Id')
            rel_type = rel.get('Type', '')
            target = rel.get('Target', '')
            
            if ('hyperlink' in rel_type.lower() or 
                'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink' in rel_type):
                hyperlinks[rel_id] = target
            elif ('image' in rel_type.lower() or 
                  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image' in rel_type):
                images[rel_id] = target
    except (KeyError, ET.ParseError):
        pass
    
    return hyperlinks, images
