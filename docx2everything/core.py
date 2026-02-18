"""
Core processing functions for docx2everything.
"""

import zipfile
from .converters.text_converter import convert_to_text
from .converters.markdown_converter import convert_to_markdown


def process(docx, img_dir=None):
    """
    Extract plain text from a DOCX file.
    
    Args:
        docx: Path to the DOCX file
        img_dir: Optional directory to extract images
    
    Returns:
        Plain text string
    """
    zipf = zipfile.ZipFile(docx)
    filelist = zipf.namelist()
    
    text = convert_to_text(zipf, filelist, img_dir)
    
    zipf.close()
    return text


def process_to_markdown(docx, img_dir=None):
    """
    Convert a DOCX file to markdown format, preserving structure like tables,
    lists, headings, formatting, links, and images.
    
    Args:
        docx: Path to the DOCX file
        img_dir: Optional directory to extract images
    
    Returns:
        Markdown string
    """
    zipf = zipfile.ZipFile(docx)
    filelist = zipf.namelist()
    
    markdown = convert_to_markdown(zipf, filelist, img_dir)
    
    zipf.close()
    return markdown
