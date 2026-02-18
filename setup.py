"""
Setup configuration for docx2everything package.
"""

import glob
from setuptools import setup

# Get all of the scripts
scripts = glob.glob('bin/*')

setup(
    name='docx2everything',
    packages=['docx2everything', 'docx2everything.utils', 
              'docx2everything.parsers', 'docx2everything.converters'],
    version='1.1.0',
    description='A pure python-based utility to extract and convert DOCX files '
                'to various formats including plain text and markdown.',
    author='sudipnext',
    maintainer='sudipnext',
    license='MIT',
    keywords=['python', 'docx', 'text', 'markdown', 'convert', 'extract'],
    scripts=scripts,
    classifiers=[
        'Development Status :: 4 - Beta',
        'Intended Audience :: Developers',
        'Topic :: Text Processing :: Markup',
        'License :: OSI Approved :: MIT License',
        'Programming Language :: Python :: 3',
    ],
    python_requires='>=3.6',
)
