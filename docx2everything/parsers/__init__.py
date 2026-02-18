"""Parsers for DOCX XML structures."""

from .relationship_parser import parse_relationships
from .numbering_parser import parse_numbering_xml
from .footnote_parser import parse_footnotes_xml, parse_endnotes_xml
from .comment_parser import parse_comments_xml
from .style_parser import parse_styles_xml
from .chart_parser import parse_all_charts

__all__ = [
    'parse_relationships',
    'parse_numbering_xml',
    'parse_footnotes_xml',
    'parse_endnotes_xml',
    'parse_comments_xml',
    'parse_styles_xml',
    'parse_all_charts'
]
