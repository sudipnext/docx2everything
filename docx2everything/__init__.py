"""
docx2everything - A pure Python utility to extract and convert DOCX files
to various formats including plain text and markdown.

Maintainer: sudipnext
"""

from .core import process, process_to_markdown
from .cli import process_args

__version__ = '1.1.0'
__author__ = 'sudipnext'

__all__ = ['process', 'process_to_markdown', 'process_args']
