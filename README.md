# Office-DDE-Payloads

## Overview
Collection of scripts and templates to generate Word documents embedded with the DDE, macro-less command execution technique described 
by [@_staaldraad](https://twitter.com/_staaldraad) and [@0x5A1F](https://twitter.com/Saif_Sherei) (blog post link in [References](#references) 
section below). Intended for use during sanctioned red team engagements and/or phishing campaigns.

Obfuscation and evasion techniques inspired by [@_staaldraad](https://twitter.com/_staaldraad) (blog post link in [References](#references) setion below).

## Usage
Insert a simple (unobfuscated) DDE command string into the payload document:

    python ddeword.py

Insert an obfuscated DDE command string by way of the {QUOTE} field code technique into the payload document:

    python ddeword.py --obfuscate
**NOTE:** *The obfuscated payload will elicit three sets of prompts to the user instead of one. Furthermore, the first set of prompts after the "update links" prompt does not update properly and cannot execute the DDE element. The second and third prompts will trigger the payload.*

Both forms of usage will generate two Word documents:

*out/template-final.docx*
- The webSettings are configured to pull the DDE element from payload-final.docx or   payload-obfuscated-final.docx, which is hosted by a server specified by the user. 

*out/payload-final.docx* (not obfuscated)  
*out/payload-obfuscated-final.docx* (obfuscated)
- Contains user-provided DDE payload/command string. Hosted by user-controlled    server (URL provided by user and baked into template-final.docx).

## Delivery
By default, the user then has two standard methods of payload delivery, described below:

1. Host the payload Word document on a user-controlled server (pointed to by the provided URL). Customize/stylize the template document and send directly to the desired target(s). This will trigger a remote reference to the payload document, ultimately pulling and executing the DDE command string.

2. Customize/Stylize the payload document and send directly to the desired target(s).

**NOTE:** *Be sure to remove personal/identifying information from the documents before hosting and sending to target (e.g., by way of the File --> Inspect Document functionality).*

## To-Do
- The obfuscated payload, by design, elicits three sets of prompts that are displayed to the user. For some reason, the first set of prompts don't display the provided DDE command string, but rather "!Unexpected end of formula". Need to devise fix for this.
- Add more obfuscation techniques
- Create option for user to choose prepackaged DDE command strings

## References
https://sensepost.com/blog/2017/macro-less-code-exec-in-msword/  
https://staaldraad.github.io/2017/10/23/msword-field-codes/

## Additional Thanks
[@ryHanson](https://twitter.com/ryhanson)  
[@SecuritySift](https://twitter.com/securitysift)  
[@GossiTheDog](https://twitter.com/gossithedog)
