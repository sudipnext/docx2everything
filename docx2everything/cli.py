"""
Command-line interface for docx2everything.
"""

import argparse
import os
import sys
from .core import process, process_to_markdown


def process_args():
    """
    Parse command-line arguments.
    
    Returns:
        argparse.Namespace: Parsed arguments
    """
    parser = argparse.ArgumentParser(
        description='A pure python-based utility to extract and convert '
                    'DOCX files to various formats (text, markdown).'
    )
    parser.add_argument("docx", help="path of the docx file")
    parser.add_argument(
        '-i', '--img_dir',
        help='path of directory to extract images'
    )
    parser.add_argument(
        '-m', '--markdown',
        action='store_true',
        help='output in markdown format instead of plain text'
    )

    args = parser.parse_args()

    if not os.path.exists(args.docx):
        print('File {} does not exist.'.format(args.docx))
        sys.exit(1)

    if args.img_dir is not None:
        if not os.path.exists(args.img_dir):
            try:
                os.makedirs(args.img_dir)
            except OSError:
                print("Unable to create img_dir {}".format(args.img_dir))
                sys.exit(1)
    
    return args


def main():
    """
    Main entry point for CLI.
    """
    args = process_args()
    if args.markdown:
        text = process_to_markdown(args.docx, args.img_dir)
    else:
        text = process(args.docx, args.img_dir)
    output = getattr(sys.stdout, 'buffer', sys.stdout)
    output.write(text.encode('utf-8'))
