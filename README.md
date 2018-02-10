# Office-DDE-Payloads

## Overview
Collection of scripts and templates to generate Word and Excel documents embedded with the DDE, macro-less command execution technique described 
by [@_staaldraad](https://twitter.com/_staaldraad) and [@0x5A1F](https://twitter.com/Saif_Sherei) (blog post link in [References](#references) 
section below). Intended for use during sanctioned red team engagements and/or phishing campaigns.

Word DDE obfuscation and evasion techniques inspired by [@_staaldraad](https://twitter.com/_staaldraad) (blog post link in [References](#references) section below).

**NOTE:** *Be sure to remove personal/identifying information from the documents before hosting and sending to target (e.g., by way of the File --> Inspect Document functionality). This has already been applied to the docs in 'templates', but it's always good practice to confirm.*

## Install Dependencies
    pip install -r requirements.txt

## Usage (Excel)
Insert a simple (unobfuscated) DDE command string into the Excel payload document:

    python ddeexcel.py

This will generate one Excel document:

*out/payload-final.xlsx*
- Contains user-provided DDE payload/command string. 

### Delivery
By default, the user then has one standard methods of payload delivery, described below:

1. Customize/Stylize the Excel payload document and send directly to the desired target(s).

## Usage (Word)
Insert a simple (unobfuscated) DDE command string into the Word payload document:

    python ddeword.py

Insert an obfuscated DDE command string by way of the {QUOTE} field code technique into the Word payload document:

    python ddeword.py --obfuscate

**NOTE:** *When leveraging the template doc to link/reference the payload doc (by way of the --obfuscate trigger), do not use the DDEAUTO element in place of the DDE element. This will (1) elicit three sets of prompts to the user and (2) disrupt/corrupt the first set of prompts to display "!Unexpected end of formula" and fail the execution of the DDE command string. For some reason, DDEAUTO does not play well with the updateFields value in word/settings.xml.*

Both forms of usage will generate two Word documents:

*out/template-final.docx*
- The webSettings are configured to pull the DDE element from payload-final.docx or   payload-obfuscated-final.docx, which is hosted by a server specified by the user. 

*out/payload-final.docx* (not obfuscated)  
*out/payload-obfuscated-final.docx* (obfuscated)
- Contains user-provided DDE payload/command string. Hosted by user-controlled    server (URL provided by user and baked into template-final.docx).

### Delivery
By default, the user then has two standard methods of payload delivery, described below:

1. Host the payload Word document on a user-controlled server (pointed to by the provided URL). Customize/stylize the template document and send directly to the desired target(s). This will trigger a remote reference to the payload document, ultimately pulling and executing the DDE command string.

2. Customize/Stylize the Word payload document and send directly to the desired target(s).

## To-Do
- Add a script to implement @enigma0x3's embedding Excel spreadsheet in OneNote technique
- Add option to include various conditional statements intended to increase operational security (e.g., sandbox evasion)
- Research obfuscation/evasion techniques for both Excel and Outlook
- Add module for Outlook DDE payload generation
- Create option for user to choose prepackaged DDE command strings

## References
https://www.contextis.com/blog/comma-separated-vulnerabilities  
http://www.exploresecurity.com/from-csv-to-cmd-to-qwerty/  
https://sensepost.com/blog/2016/powershell-c-sharp-and-dde-the-power-within/  
https://sensepost.com/blog/2017/macro-less-code-exec-in-msword/  
https://staaldraad.github.io/2017/10/23/msword-field-codes/  
https://posts.specterops.io/reviving-dde-using-onenote-and-excel-for-code-execution-d7226864caee  

## Additional Thanks
[@ryHanson](https://twitter.com/ryhanson)  
[@SecuritySift](https://twitter.com/securitysift)  
[@GossiTheDog](https://twitter.com/gossithedog)  
[@enigma0x3](https://twitter.com/enigma0x3)  
[@albinowax](https://twitter.com/albinowax)  
[@exploresecurity](https://twitter.com/exploresecurity)  
[@decode141](https://twitter.com/decode141)  
