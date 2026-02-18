"""
File utility functions for handling DOCX files and image extraction.
"""

import os
import zipfile


def extract_images(zipf, filelist, img_dir):
    """
    Extract images from DOCX file to specified directory.
    
    Args:
        zipf: ZipFile object of the DOCX file
        filelist: List of files in the ZIP archive
        img_dir: Target directory for image extraction
    
    Returns:
        None (images are written to disk)
    """
    if img_dir is None:
        return
    
    for fname in filelist:
        _, extension = os.path.splitext(fname)
        if extension.lower() in [".jpg", ".jpeg", ".png", ".bmp", ".gif"]:
            dst_fname = os.path.join(img_dir, os.path.basename(fname))
            try:
                with open(dst_fname, "wb") as dst_f:
                    dst_f.write(zipf.read(fname))
            except (OSError, IOError):
                pass
