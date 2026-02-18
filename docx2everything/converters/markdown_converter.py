"""
Markdown converter for DOCX files.
Converts DOCX structures (tables, lists, headings, formatting) to markdown.
"""

import os
import re
import xml.etree.ElementTree as ET
from ..utils.xml_utils import qn, NSMAP
from ..parsers.relationship_parser import parse_relationships
from ..parsers.numbering_parser import parse_numbering_xml
from ..utils.file_utils import extract_images


def get_heading_level(pStyle_val):
    """
    Maps paragraph style to markdown heading level.
    
    Args:
        pStyle_val: Paragraph style value
    
    Returns:
        int: Heading level (1-6) or None if not a heading
    """
    if not pStyle_val:
        return None
    
    pStyle_lower = pStyle_val.lower()
    if 'title' in pStyle_lower:
        return 1
    elif 'heading1' in pStyle_lower or 'heading 1' in pStyle_lower:
        return 1
    elif 'heading2' in pStyle_lower or 'heading 2' in pStyle_lower:
        return 2
    elif 'heading3' in pStyle_lower or 'heading 3' in pStyle_lower:
        return 3
    elif 'heading4' in pStyle_lower or 'heading 4' in pStyle_lower:
        return 4
    elif 'heading5' in pStyle_lower or 'heading 5' in pStyle_lower:
        return 5
    elif 'heading6' in pStyle_lower or 'heading 6' in pStyle_lower:
        return 6
    return None


def parse_run_to_markdown(r_elem, hyperlinks=None, images=None, img_dir=None, zipf=None, link_url=None):
    """
    Converts a text run (<w:r>) to markdown with formatting.
    
    Args:
        r_elem: XML element representing a text run
        hyperlinks: Dict mapping relationship IDs to URLs
        images: Dict mapping relationship IDs to image paths
        img_dir: Directory for extracted images
        zipf: ZipFile object
        link_url: Optional URL if this run is part of a hyperlink
    
    Returns:
        Markdown string with formatting
    """
    text = ''
    rPr = r_elem.find(qn('w:rPr'))
    
    # Extract text from runs
    for t_elem in r_elem.findall(qn('w:t')):
        if t_elem.text:
            text += t_elem.text
    
    # Handle tabs and breaks
    for tab in r_elem.findall(qn('w:tab')):
        text += '    '  # Convert tab to 4 spaces
    for br in r_elem.findall(qn('w:br')):
        text += '\n'
    
    if not text:
        return ''
    
    # Apply formatting (check all formatting first, then apply appropriately)
    if rPr is not None:
        is_bold = rPr.find(qn('w:b')) is not None
        is_italic = rPr.find(qn('w:i')) is not None
        is_strike = (rPr.find(qn('w:strike')) is not None or 
                     rPr.find(qn('w:delText')) is not None)
        
        # Apply formatting in correct order (strikethrough, then bold/italic)
        if is_strike:
            text = '~~' + text + '~~'
        
        # Bold and italic together
        if is_bold and is_italic:
            text = '***' + text + '***'
        elif is_bold:
            text = '**' + text + '**'
        elif is_italic:
            text = '*' + text + '*'
    
    # Wrap in hyperlink if provided
    if link_url:
        text = '[' + text + '](' + link_url + ')'
    
    return text


def parse_paragraph_to_markdown(p_elem, numbering_info=None, hyperlinks=None, images=None, img_dir=None, zipf=None):
    """
    Converts a paragraph (<w:p>) to markdown.
    Handles headings, lists, regular paragraphs, and formatting.
    
    Args:
        p_elem: XML element representing a paragraph
        numbering_info: Dict mapping numId to list information
        hyperlinks: Dict mapping relationship IDs to URLs
        images: Dict mapping relationship IDs to image paths
        img_dir: Directory for extracted images
        zipf: ZipFile object
    
    Returns:
        Markdown string
    """
    pPr = p_elem.find(qn('w:pPr'))
    
    # Check for heading
    heading_level = None
    if pPr is not None:
        pStyle = pPr.find(qn('w:pStyle'))
        if pStyle is not None:
            style_val = pStyle.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val')
            heading_level = get_heading_level(style_val)
    
    # Check for list item
    is_list_item = False
    list_info = None
    if pPr is not None:
        numPr = pPr.find(qn('w:numPr'))
        if numPr is not None and numbering_info:
            ilvl_elem = numPr.find(qn('w:ilvl'))
            numId_elem = numPr.find(qn('w:numId'))
            if ilvl_elem is not None and numId_elem is not None:
                ilvl = int(ilvl_elem.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', 0))
                numId = numId_elem.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val')
                if numId in numbering_info:
                    is_list_item = True
                    list_info = {
                        'ilvl': ilvl,
                        'numId': numId,
                        'list_type': numbering_info[numId].get('list_type', 'bullet'),
                        'num_format': numbering_info[numId].get('num_format', 'decimal')
                    }
    
    # Extract text from runs
    para_text = ''
    
    # Process all child elements in order (runs and hyperlinks)
    for child in p_elem:
        if child.tag == qn('w:r'):
            # Regular run
            para_text += parse_run_to_markdown(child, hyperlinks, images, img_dir, zipf)
        elif child.tag == qn('w:hyperlink'):
            # Hyperlink - process runs inside it
            rel_id = child.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
            link_url = '#'
            if rel_id and hyperlinks:
                link_url = hyperlinks.get(rel_id, '#')
            
            link_text = ''
            for r in child.findall(qn('w:r')):
                link_text += parse_run_to_markdown(r, hyperlinks, images, img_dir, zipf, link_url=None)
            
            if link_text:
                para_text += '[' + link_text + '](' + link_url + ')'
    
    # Handle images
    for drawing in p_elem.findall('.//' + qn('w:drawing')):
        # Try to extract image relationship
        blip = drawing.find('.//{http://schemas.openxmlformats.org/drawingml/2006/main}blip')
        if blip is not None:
            rel_id = blip.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed')
            if not rel_id:
                rel_id = blip.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}link')
            
            if rel_id and images:
                img_path = images.get(rel_id, '')
                if img_dir and img_path:
                    # Use relative path from img_dir
                    img_filename = os.path.basename(img_path)
                    img_md_path = os.path.join(img_dir, img_filename) if img_dir else img_path
                    para_text += '\n![' + img_filename + '](' + img_md_path + ')\n'
                elif img_path:
                    para_text += '\n![' + os.path.basename(img_path) + '](' + img_path + ')\n'
    
    para_text = para_text.strip()
    
    if not para_text and not is_list_item:
        return ''
    
    # Format based on type
    if heading_level:
        return '#' * heading_level + ' ' + para_text
    elif is_list_item:
        indent = '  ' * list_info['ilvl']
        if list_info['list_type'] == 'bullet':
            return indent + '- ' + para_text
        else:
            # For numbered lists, we'll use a simple counter
            # In a full implementation, we'd track the actual number
            return indent + '1. ' + para_text
    else:
        return para_text


def parse_table_to_markdown(tbl_elem, hyperlinks=None, images=None, img_dir=None, zipf=None):
    """
    Converts a table (<w:tbl>) to markdown table syntax.
    
    Args:
        tbl_elem: XML element representing a table
        hyperlinks: Dict mapping relationship IDs to URLs
        images: Dict mapping relationship IDs to image paths
        img_dir: Directory for extracted images
        zipf: ZipFile object
    
    Returns:
        Markdown table string
    """
    rows = tbl_elem.findall(qn('w:tr'))
    if not rows:
        return ''
    
    markdown_rows = []
    num_cols = 0
    
    # First pass: determine number of columns and extract all rows
    for row in rows:
        cells = row.findall(qn('w:tc'))
        row_data = []
        
        for cell in cells:
            # Check for gridSpan (merged cells)
            tcPr = cell.find(qn('w:tcPr'))
            grid_span = 1
            if tcPr is not None:
                gridSpan_elem = tcPr.find(qn('w:gridSpan'))
                if gridSpan_elem is not None:
                    grid_span = int(gridSpan_elem.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', 1))
            
            # Extract cell text
            cell_text = ''
            for p in cell.findall(qn('w:p')):
                p_text = parse_paragraph_to_markdown(p, None, hyperlinks, images, img_dir, zipf)
                if p_text:
                    cell_text += p_text + ' '
            
            cell_text = cell_text.strip().replace('\n', ' ').replace('|', '\\|')
            
            # Add merged cells
            row_data.append(cell_text)
            for _ in range(grid_span - 1):
                row_data.append('')  # Empty cells for merged columns
        
        if row_data:
            markdown_rows.append(row_data)
            num_cols = max(num_cols, len(row_data))
    
    if not markdown_rows:
        return ''
    
    # Normalize column count
    for row in markdown_rows:
        while len(row) < num_cols:
            row.append('')
    
    # Build markdown table
    md_table = []
    
    # Header row (assume first row is header, or detect via style)
    header_row = markdown_rows[0]
    md_table.append('| ' + ' | '.join(header_row) + ' |')
    
    # Separator row
    md_table.append('| ' + ' | '.join(['---'] * num_cols) + ' |')
    
    # Data rows
    for row in markdown_rows[1:]:
        md_table.append('| ' + ' | '.join(row) + ' |')
    
    return '\n'.join(md_table)


def parse_body_to_markdown(root, numbering_info=None, hyperlinks=None, images=None, img_dir=None, zipf=None):
    """
    Main parser that traverses document body and converts elements to markdown.
    
    Args:
        root: XML root element
        numbering_info: Dict mapping numId to list information
        hyperlinks: Dict mapping relationship IDs to URLs
        images: Dict mapping relationship IDs to image paths
        img_dir: Directory for extracted images
        zipf: ZipFile object
    
    Returns:
        Markdown string
    """
    markdown_parts = []
    body = root.find(qn('w:body'))
    
    if body is None:
        return ''
    
    for elem in body:
        if elem.tag == qn('w:p'):
            # Paragraph
            para_md = parse_paragraph_to_markdown(elem, numbering_info, hyperlinks, images, img_dir, zipf)
            if para_md:
                markdown_parts.append(para_md)
        elif elem.tag == qn('w:tbl'):
            # Table
            table_md = parse_table_to_markdown(elem, hyperlinks, images, img_dir, zipf)
            if table_md:
                markdown_parts.append(table_md)
                markdown_parts.append('')  # Add blank line after table
    
    return '\n\n'.join(markdown_parts)


def convert_to_markdown(zipf, filelist, img_dir=None):
    """
    Converts DOCX file to markdown format.
    
    Args:
        zipf: ZipFile object of the DOCX file
        filelist: List of files in the ZIP archive
        img_dir: Optional directory to extract images
    
    Returns:
        Markdown string
    """
    markdown_parts = []
    
    # Parse relationships
    hyperlinks, images = parse_relationships(zipf)
    
    # Parse numbering information
    numbering_info = parse_numbering_xml(zipf)
    
    # Extract images if needed
    if img_dir is not None:
        extract_images(zipf, filelist, img_dir)
    
    # Process headers
    header_xmls = 'word/header[0-9]*.xml'
    for fname in filelist:
        if re.match(header_xmls, fname):
            header_xml = zipf.read(fname)
            header_root = ET.fromstring(header_xml)
            header_md = parse_body_to_markdown(header_root, numbering_info, hyperlinks, images, img_dir, zipf)
            if header_md:
                markdown_parts.append(header_md)
    
    # Process main document
    doc_xml = 'word/document.xml'
    doc_xml_content = zipf.read(doc_xml)
    doc_root = ET.fromstring(doc_xml_content)
    doc_md = parse_body_to_markdown(doc_root, numbering_info, hyperlinks, images, img_dir, zipf)
    if doc_md:
        markdown_parts.append(doc_md)
    
    # Process footers
    footer_xmls = 'word/footer[0-9]*.xml'
    for fname in filelist:
        if re.match(footer_xmls, fname):
            footer_xml = zipf.read(fname)
            footer_root = ET.fromstring(footer_xml)
            footer_md = parse_body_to_markdown(footer_root, numbering_info, hyperlinks, images, img_dir, zipf)
            if footer_md:
                markdown_parts.append(footer_md)
    
    result = '\n\n'.join(markdown_parts)
    return result.strip()
