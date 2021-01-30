#!/usr/bin/env python3

# A very simple Python 3 script to extract all TEXT elements from all
# PARAGRAPH elements of elderly KWord documents. This was bashed
# together in an hour and I'd never used the zipfile, gzip, tarfile
# or lxml libraries before, and it's not meant to be pretty. I just
# had a hundred old documents I wanted to grab the words from.

# If it helps, my vintage .kwd documents turned out to be zipped (both
# old-school Zip, i.e. PK, and gzipped tarballs; I had a mixture of
# both) archives whose maindoc.xml file was the actual document, and
# the structure seemed simple enough. I just wrote some XSL to extract
# the text and threw together this script so I could do the extraction
# easily.

# THIS IS ENTIRELY UNSUPPORTED AND PROBABLY VERY BADLY WRITTEN, but
# if you need to get text out of an old KWord file, it might be worth
# a try for you, so here it is.

# You'll probably need to install lxml with pip.

import argparse
import sys
import os
import zipfile
import gzip
import tarfile
import lxml.etree as ET

# This XSL is designed to cope with a couple of different formats I found
# in the maindoc.xml files in my .kwd documents. Basically there was a
# nicely-namespaced format and a non-namespaced format that I'm guessing
# was an earlier version. Either way, it just finds all the TEXT elements
# inside all PARAGRAPH elements and prints them out.
xsl="""<?xml version="1.0"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
		xmlns:kword="http://www.koffice.org/DTD/kword">

  <xsl:output method="text"/>

  <xsl:template match="/">
    <xsl:apply-templates select="//kword:PARAGRAPH|//PARAGRAPH" />
  </xsl:template>

  <xsl:template match="PARAGRAPH|kword:PARAGRAPH">
    <xsl:apply-templates select=".//kword:TEXT|.//TEXT" />
  </xsl:template>

  <xsl:template match="TEXT|kword:TEXT">
    <xsl:value-of select="." />
    <xsl:text>
    </xsl:text>
  </xsl:template>

</xsl:stylesheet>
"""

parser = argparse.ArgumentParser(
    description='Extract all text from paragraphs inside an old-school KOffice KWord document.'
)

parser.add_argument("kword", help="An elderly KWord document.")
parser.add_argument("text", help="An output text file (that will be overwritten if exists.)")

args = parser.parse_args()
infile_name = os.path.abspath(args.kword)
outfile_name = os.path.abspath(args.text)

# This is terribly quick and dirty and does a lot of things in memory that could probably
# be better streamed. On the other hand, I just want to get this done, and my KWord files
# are small because they're from fifteen years ago and my Mac has 32GB, and I'm only
# writing this to use once...
transform = ET.XSLT(ET.ElementTree(ET.fromstring(xsl)))

# with zipfile.ZipFile(infile_name, 'r') as infile:
# xmldata = ET.ElementTree(ET.fromstring(infile.read('maindoc.xml')))
infile = None
try:
    infile = tarfile.open(infile_name, "r:gz")
    xmldata = ET.ElementTree(ET.fromstring(infile.extractfile('maindoc.xml').read()))
except (gzip.BadGzipFile, tarfile.ReadError):
    # It might not be a tarball
    pass

if infile is None:
    # If it wasn't a tarball, maybe it was a PKZip
    infile = zipfile.ZipFile(infile_name, 'r')
    xmldata = ET.ElementTree(ET.fromstring(infile.read('maindoc.xml')))

result = transform(xmldata)
with open(outfile_name, "w") as outfile:
    outfile.write(str(result))

infile.close()

