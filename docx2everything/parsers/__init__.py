"""Parsers for DOCX XML structures."""

from .relationship_parser import parse_relationships
from .numbering_parser import parse_numbering_xml

__all__ = ['parse_relationships', 'parse_numbering_xml']
