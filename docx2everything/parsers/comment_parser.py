"""
Parser for DOCX comments.
"""

import xml.etree.ElementTree as ET
from ..utils.xml_utils import qn, NSMAP


def parse_comments_xml(zipf):
    """
    Parses comments.xml to extract comment content.
    
    Args:
        zipf: ZipFile object of the DOCX file
    
    Returns:
        dict: Mapping of comment ID to comment text and author
    """
    comments = {}
    
    try:
        comments_xml = zipf.read('word/comments.xml')
        root = ET.fromstring(comments_xml)
        
        for comment in root.findall('.//{' + NSMAP['w'] + '}comment'):
            comment_id = comment.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}id')
            author = comment.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}author', 'Unknown')
            date = comment.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}date', '')
            
            # Extract text from paragraphs in comment
            comment_text = ''
            for para in comment.findall('.//{' + NSMAP['w'] + '}p'):
                for run in para.findall('.//{' + NSMAP['w'] + '}r'):
                    for t in run.findall('.//{' + NSMAP['w'] + '}t'):
                        if t.text:
                            comment_text += t.text
                    for br in run.findall('.//{' + NSMAP['w'] + '}br'):
                        comment_text += '\n'
                comment_text += '\n'
            
            if comment_text.strip():
                comments[comment_id] = {
                    'text': comment_text.strip(),
                    'author': author,
                    'date': date
                }
    except (KeyError, ET.ParseError):
        # If comments.xml doesn't exist or can't be parsed
        pass
    
    return comments
