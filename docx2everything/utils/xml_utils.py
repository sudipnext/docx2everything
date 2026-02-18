"""
XML utility functions for parsing DOCX files.
"""

NSMAP = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
}


def qn(tag):
    """
    Stands for 'qualified name', a utility function to turn a namespace
    prefixed tag name into a Clark-notation qualified tag name.
    
    Example: ``qn('w:p')`` returns ``'{http://schemas.../main}p'``
    
    Args:
        tag: A namespace-prefixed tag name (e.g., 'w:p', 'w:t')
    
    Returns:
        A Clark-notation qualified tag name
    """
    prefix, tagroot = tag.split(':')
    uri = NSMAP[prefix]
    return '{{{}}}{}'.format(uri, tagroot)
