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
from ..parsers.footnote_parser import parse_footnotes_xml, parse_endnotes_xml
from ..parsers.comment_parser import parse_comments_xml
from ..parsers.style_parser import parse_styles_xml
from ..parsers.chart_parser import parse_all_charts
from ..utils.file_utils import extract_images


def get_heading_level(pStyle_val, styles_info=None):
    """
    Maps paragraph style to markdown heading level.
    Uses styles.xml information if available for better custom style detection.
    
    Args:
        pStyle_val: Paragraph style value
        styles_info: Dict mapping style IDs to style information (from styles.xml)
    
    Returns:
        int: Heading level (1-6) or None if not a heading
    """
    if not pStyle_val:
        return None
    
    # First check styles.xml if available
    if styles_info and pStyle_val in styles_info:
        style_info = styles_info[pStyle_val]
        if style_info.get('is_heading'):
            level = style_info.get('heading_level')
            if level:
                return level
    
    # Fallback to pattern matching
    pStyle_lower = pStyle_val.lower()
    if 'title' in pStyle_lower:
        return 1
    elif 'heading1' in pStyle_lower or 'heading 1' in pStyle_lower or 'h1' == pStyle_lower:
        return 1
    elif 'heading2' in pStyle_lower or 'heading 2' in pStyle_lower or 'h2' == pStyle_lower:
        return 2
    elif 'heading3' in pStyle_lower or 'heading 3' in pStyle_lower or 'h3' == pStyle_lower:
        return 3
    elif 'heading4' in pStyle_lower or 'heading 4' in pStyle_lower or 'h4' == pStyle_lower:
        return 4
    elif 'heading5' in pStyle_lower or 'heading 5' in pStyle_lower or 'h5' == pStyle_lower:
        return 5
    elif 'heading6' in pStyle_lower or 'heading 6' in pStyle_lower or 'h6' == pStyle_lower:
        return 6
    
    # Check if based_on style is a heading (recursive check)
    if styles_info and pStyle_val in styles_info:
        based_on = styles_info[pStyle_val].get('based_on')
        if based_on:
            return get_heading_level(based_on, styles_info)
    
    return None


def parse_run_to_markdown(r_elem, hyperlinks=None, images=None, img_dir=None, zipf=None, link_url=None, footnotes=None, endnotes=None):
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
    
    # Handle footnote references
    for footnote_ref in r_elem.findall(qn('w:footnoteReference')):
        footnote_id = footnote_ref.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}id')
        if footnotes and footnote_id in footnotes:
            text += f'[^{footnote_id}]'
        else:
            text += f'[^{footnote_id}]'
    
    # Handle endnote references
    for endnote_ref in r_elem.findall(qn('w:endnoteReference')):
        endnote_id = endnote_ref.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}id')
        if endnotes and endnote_id in endnotes:
            text += f'[^{endnote_id}]'
        else:
            text += f'[^{endnote_id}]'
    
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


def parse_paragraph_to_markdown(p_elem, numbering_info=None, hyperlinks=None, images=None, img_dir=None, zipf=None, footnotes=None, endnotes=None, comments=None, list_counters=None, styles_info=None, charts=None):
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
    
    # Check for page break
    has_page_break = False
    if pPr is not None:
        page_break_before = pPr.find(qn('w:pageBreakBefore'))
        if page_break_before is not None:
            has_page_break = True
    
    # Check for section break
    has_section_break = False
    sectPr = p_elem.find(qn('w:sectPr'))
    if sectPr is not None:
        has_section_break = True
    
    # Check for heading
    heading_level = None
    if pPr is not None:
        pStyle = pPr.find(qn('w:pStyle'))
        if pStyle is not None:
            style_val = pStyle.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val')
            heading_level = get_heading_level(style_val, styles_info)
    
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
                    
                    # Track list counters for numbered lists
                    if list_info['list_type'] == 'number' and list_counters is not None:
                        list_key = f"{numId}_{ilvl}"
                        if list_key not in list_counters:
                            list_counters[list_key] = 0
                        list_counters[list_key] += 1
                        list_info['counter'] = list_counters[list_key]
    
    # Extract text from runs
    para_text = ''
    
    # Process all child elements in order (runs and hyperlinks)
    for child in p_elem:
        if child.tag == qn('w:r'):
            # Regular run
            para_text += parse_run_to_markdown(child, hyperlinks, images, img_dir, zipf, footnotes=footnotes, endnotes=endnotes)
        elif child.tag == qn('w:hyperlink'):
            # Hyperlink - process runs inside it
            rel_id = child.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
            link_url = '#'
            if rel_id and hyperlinks:
                link_url = hyperlinks.get(rel_id, '#')
            
            link_text = ''
            for r in child.findall(qn('w:r')):
                link_text += parse_run_to_markdown(r, hyperlinks, images, img_dir, zipf, link_url=None, footnotes=footnotes, endnotes=endnotes)
            
            if link_text:
                para_text += '[' + link_text + '](' + link_url + ')'
        elif child.tag == qn('w:commentRangeStart'):
            # Comment start - we'll handle this with commentRangeEnd
            pass
        elif child.tag == qn('w:commentRangeEnd'):
            # Comment end - extract comment
            comment_id = child.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}id')
            if comments and comment_id in comments:
                comment_data = comments[comment_id]
                para_text += f' <!-- Comment by {comment_data["author"]}: {comment_data["text"][:50]}... -->'
    
    # Handle images and charts
    for drawing in p_elem.findall('.//' + qn('w:drawing')):
        # Check for charts - charts can be in graphicFrame or directly in graphic
        chart_ref = drawing.find('.//{http://schemas.openxmlformats.org/drawingml/2006/chart}chart')
        if chart_ref is not None:
            chart_rel_id = chart_ref.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
            if chart_rel_id:
                        if charts and chart_rel_id in charts:
                            chart_info = charts[chart_rel_id]
                            chart_title = chart_info.get('title', 'Chart')
                            chart_type = chart_info.get('chart_type', 'Chart')
                            data_points = chart_info.get('data_points', [])
                            
                            para_text += f'\n\n<!-- Chart: {chart_title}'
                            if chart_type:
                                para_text += f' ({chart_type})'
                            para_text += ' -->\n'
                            para_text += f'*[Chart: {chart_title}'
                            if chart_type:
                                para_text += f' ({chart_type})'
                            para_text += ']*\n'
                            
                            # Add chart data if available
                            if data_points:
                                para_text += '\n```\n'
                                para_text += 'Chart Data:\n'
                                for i, series in enumerate(data_points):
                                    series_name = series.get('series_name', f'Series {i+1}')
                                    values = series.get('values', [])
                                    categories = series.get('categories')
                                    
                                    para_text += f'\n{series_name}:\n'
                                    if categories and len(categories) == len(values):
                                        # Show as table format
                                        for cat, val in zip(categories, values):
                                            para_text += f'  {cat}: {val}\n'
                                    else:
                                        # Just show values
                                        para_text += f'  Values: {", ".join(map(str, values))}\n'
                                para_text += '```\n'
                            elif chart_info.get('has_data'):
                                para_text += '<!-- Chart contains data (embedded Excel reference) -->\n'
                        else:
                            para_text += '\n\n*[Chart (relationship ID: ' + chart_rel_id + ') - data not available]*\n'
                        continue
        
        # Check for charts in graphicFrame elements
        graphic_frame = drawing.find('.//{http://schemas.openxmlformats.org/drawingml/2006/main}graphicFrame')
        if graphic_frame is not None:
            graphic = graphic_frame.find('.//{http://schemas.openxmlformats.org/drawingml/2006/main}graphic')
            if graphic is not None:
                chart_ref = graphic.find('.//{http://schemas.openxmlformats.org/drawingml/2006/chart}chart')
                if chart_ref is not None:
                    chart_rel_id = chart_ref.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
                    if chart_rel_id:
                        if charts and chart_rel_id in charts:
                            chart_info = charts[chart_rel_id]
                            chart_title = chart_info.get('title', 'Chart')
                            chart_type = chart_info.get('chart_type', 'Chart')
                            data_points = chart_info.get('data_points', [])
                            
                            para_text += f'\n\n<!-- Chart: {chart_title}'
                            if chart_type:
                                para_text += f' ({chart_type})'
                            para_text += ' -->\n'
                            para_text += f'*[Chart: {chart_title}'
                            if chart_type:
                                para_text += f' ({chart_type})'
                            para_text += ']*\n'
                            
                            # Add chart data if available
                            if data_points:
                                para_text += '\n```\n'
                                para_text += 'Chart Data:\n'
                                for i, series in enumerate(data_points):
                                    series_name = series.get('series_name', f'Series {i+1}')
                                    values = series.get('values', [])
                                    categories = series.get('categories')
                                    
                                    para_text += f'\n{series_name}:\n'
                                    if categories and len(categories) == len(values):
                                        # Show as table format
                                        for cat, val in zip(categories, values):
                                            para_text += f'  {cat}: {val}\n'
                                    else:
                                        # Just show values
                                        para_text += f'  Values: {", ".join(map(str, values))}\n'
                                para_text += '```\n'
                            elif chart_info.get('has_data'):
                                para_text += '<!-- Chart contains data (embedded Excel reference) -->\n'
                        else:
                            para_text += '\n\n*[Chart (relationship ID: ' + chart_rel_id + ') - data not available]*\n'
                        continue
        
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
    
    # Add page/section break markers
    prefix = ''
    if has_page_break:
        prefix = '<!-- Page Break -->\n\n'
    elif has_section_break:
        prefix = '<!-- Section Break -->\n\n'
    
    if not para_text and not is_list_item:
        return prefix if prefix else ''
    
    # Format based on type
    if heading_level:
        return prefix + '#' * heading_level + ' ' + para_text
    elif is_list_item:
        indent = '  ' * list_info['ilvl']
        if list_info['list_type'] == 'bullet':
            return prefix + indent + '- ' + para_text
        else:
            # Use tracked counter for numbered lists
            counter = list_info.get('counter', 1)
            return prefix + indent + f'{counter}. ' + para_text
    else:
        return prefix + para_text


def parse_table_to_markdown(tbl_elem, hyperlinks=None, images=None, img_dir=None, zipf=None, footnotes=None, endnotes=None, styles_info=None):
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
    col_alignments = []  # Track column alignments
    for row_idx, row in enumerate(rows):
        cells = row.findall(qn('w:tc'))
        row_data = []
        row_alignments = []
        
        for cell in cells:
            # Check for gridSpan (merged cells)
            tcPr = cell.find(qn('w:tcPr'))
            grid_span = 1
            cell_alignment = 'left'  # Default alignment
            
            if tcPr is not None:
                gridSpan_elem = tcPr.find(qn('w:gridSpan'))
                if gridSpan_elem is not None:
                    grid_span = int(gridSpan_elem.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', 1))
                
                # Check for cell alignment
                jc_elem = tcPr.find(qn('w:jc'))
                if jc_elem is not None:
                    jc_val = jc_elem.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', 'left')
                    if jc_val == 'center':
                        cell_alignment = 'center'
                    elif jc_val == 'right':
                        cell_alignment = 'right'
                    elif jc_val == 'both' or jc_val == 'distribute':
                        cell_alignment = 'justify'
            
            # Extract cell text
            cell_text = ''
            for p in cell.findall(qn('w:p')):
                p_text = parse_paragraph_to_markdown(p, None, hyperlinks, images, img_dir, zipf, footnotes=footnotes, endnotes=endnotes, styles_info=styles_info)
                if p_text:
                    cell_text += p_text + ' '
            
            cell_text = cell_text.strip().replace('\n', ' ').replace('|', '\\|')
            
            # Add merged cells
            row_data.append(cell_text)
            row_alignments.append(cell_alignment)
            for _ in range(grid_span - 1):
                row_data.append('')  # Empty cells for merged columns
                row_alignments.append('left')  # Default for merged cells
        
        if row_data:
            markdown_rows.append(row_data)
            num_cols = max(num_cols, len(row_data))
            
            # Track alignments (use first row's alignments as column defaults)
            if row_idx == 0:
                col_alignments = row_alignments[:]
            else:
                # Merge alignments, preferring non-left alignments
                for i, align in enumerate(row_alignments):
                    if i < len(col_alignments) and align != 'left':
                        col_alignments[i] = align
    
    if not markdown_rows:
        return ''
    
    # Normalize column count
    for row in markdown_rows:
        while len(row) < num_cols:
            row.append('')
    
    # Build markdown table with alignment hints
    md_table = []
    
    # Header row (assume first row is header, or detect via style)
    header_row = markdown_rows[0]
    md_table.append('| ' + ' | '.join(header_row) + ' |')
    
    # Separator row with alignment hints
    separator_parts = []
    for i in range(num_cols):
        align = col_alignments[i] if i < len(col_alignments) else 'left'
        if align == 'center':
            separator_parts.append(':---:')
        elif align == 'right':
            separator_parts.append('---:')
        else:
            separator_parts.append('---')
    md_table.append('| ' + ' | '.join(separator_parts) + ' |')
    
    # Data rows
    for row in markdown_rows[1:]:
        md_table.append('| ' + ' | '.join(row) + ' |')
    
    # Add HTML comment with table formatting info if non-default alignments exist
    if any(align != 'left' for align in col_alignments):
        align_info = ', '.join([f'col{i+1}:{align}' for i, align in enumerate(col_alignments) if align != 'left'])
        md_table.append(f'<!-- Table alignment: {align_info} -->')
    
    return '\n'.join(md_table)


def parse_body_to_markdown(root, numbering_info=None, hyperlinks=None, images=None, img_dir=None, zipf=None, footnotes=None, endnotes=None, comments=None, styles_info=None, charts=None):
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
    
    # Initialize list counters for tracking numbered list sequences
    list_counters = {}
    
    for elem in body:
        if elem.tag == qn('w:p'):
            # Paragraph
            para_md = parse_paragraph_to_markdown(elem, numbering_info, hyperlinks, images, img_dir, zipf, footnotes=footnotes, endnotes=endnotes, comments=comments, list_counters=list_counters, styles_info=styles_info, charts=charts)
            if para_md:
                markdown_parts.append(para_md)
        elif elem.tag == qn('w:tbl'):
            # Table
            table_md = parse_table_to_markdown(elem, hyperlinks, images, img_dir, zipf, footnotes=footnotes, endnotes=endnotes, styles_info=styles_info)
            if table_md:
                markdown_parts.append(table_md)
                markdown_parts.append('')  # Add blank line after table
    
    # Append footnotes and endnotes at the end
    footnote_parts = []
    if footnotes:
        for footnote_id, footnote_text in sorted(footnotes.items(), key=lambda x: int(x[0]) if x[0].isdigit() else 0):
            footnote_parts.append(f'[^{footnote_id}]: {footnote_text}')
    
    if endnotes:
        for endnote_id, endnote_text in sorted(endnotes.items(), key=lambda x: int(x[0]) if x[0].isdigit() else 0):
            footnote_parts.append(f'[^{endnote_id}]: {endnote_text}')
    
    if footnote_parts:
        markdown_parts.append('')  # Blank line before footnotes
        markdown_parts.extend(footnote_parts)
    
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
    try:
        hyperlinks, images = parse_relationships(zipf)
    except Exception:
        hyperlinks, images = {}, {}
    
    # Parse numbering information
    try:
        numbering_info = parse_numbering_xml(zipf)
    except Exception:
        numbering_info = {}
    
    # Parse footnotes and endnotes
    try:
        footnotes = parse_footnotes_xml(zipf)
    except Exception:
        footnotes = {}
    
    try:
        endnotes = parse_endnotes_xml(zipf)
    except Exception:
        endnotes = {}
    
    # Parse comments
    try:
        comments = parse_comments_xml(zipf)
    except Exception:
        comments = {}
    
    # Parse styles information
    try:
        styles_info = parse_styles_xml(zipf)
    except Exception:
        styles_info = {}
    
    # Parse charts
    try:
        charts = parse_all_charts(zipf)
    except Exception:
        charts = {}
    
    # Extract images if needed
    if img_dir is not None:
        try:
            extract_images(zipf, filelist, img_dir)
        except Exception:
            pass  # Continue even if image extraction fails
    
    # Process headers
    header_xmls = 'word/header[0-9]*.xml'
    for fname in filelist:
        if re.match(header_xmls, fname):
            try:
                header_xml = zipf.read(fname)
                header_root = ET.fromstring(header_xml)
                header_md = parse_body_to_markdown(header_root, numbering_info, hyperlinks, images, img_dir, zipf, footnotes=footnotes, endnotes=endnotes, comments=comments, styles_info=styles_info, charts=charts)
                if header_md:
                    markdown_parts.append(header_md)
            except Exception:
                pass  # Skip if header parsing fails
    
    # Process main document
    try:
        doc_xml = 'word/document.xml'
        doc_xml_content = zipf.read(doc_xml)
        doc_root = ET.fromstring(doc_xml_content)
        doc_md = parse_body_to_markdown(doc_root, numbering_info, hyperlinks, images, img_dir, zipf, footnotes=footnotes, endnotes=endnotes, comments=comments, styles_info=styles_info, charts=charts)
        if doc_md:
            markdown_parts.append(doc_md)
    except Exception as e:
        # If main document fails, try to extract at least basic text
        markdown_parts.append(f'<!-- Error parsing document: {str(e)} -->')
    
    # Process footers
    footer_xmls = 'word/footer[0-9]*.xml'
    for fname in filelist:
        if re.match(footer_xmls, fname):
            try:
                footer_xml = zipf.read(fname)
                footer_root = ET.fromstring(footer_xml)
                footer_md = parse_body_to_markdown(footer_root, numbering_info, hyperlinks, images, img_dir, zipf, footnotes=footnotes, endnotes=endnotes, comments=comments, styles_info=styles_info, charts=charts)
                if footer_md:
                    markdown_parts.append(footer_md)
            except Exception:
                pass  # Skip if footer parsing fails
    
    result = '\n\n'.join(markdown_parts)
    return result.strip()
