# Office-DDE-Payloads

## Overview
Collection of scripts and templates to generate Office documents embedded with the DDE, macro-less command execution technique described 
by [@_staaldraad](https://twitter.com/_staaldraad) and [@0x5A1F](https://twitter.com/Saif_Sherei) (blog post link in [References](#references) 
section below). Best if used during red team engagements and/or phishing campaigns.

Obfuscation and evasion techniques inspired by [@_staaldraad](https://twitter.com/_staaldraad) (blog post link in [References](#references) setion below).

## Usage
If you'd like to insert a simple DDE command string into the payload document, run:

    python ddeword.py

If you'd prefer to obfuscate the DDE command string by way of the {QUOTE} field code technique, run:

    python ddeword.py --obfuscate

Both forms of usage will generate two Word documents:

*template-final.docx*
- The webSettings are configured to pull the DDE element from payload-final.docx or   payload-obfuscated-final.docx, which is hosted by a server specified by the user. 

*payload-final.docx* (not obfuscated)  
*payload-obfuscated-final.docx* (obfuscated)
- Contains user-provided DDE payload/command string. Hosted by user-controlled    server (URL provided by user and baked into template-final.docx).

## Delivery
By default, the user then has two standard methods of payload delivery, described below:

1. Host the payload Office document on a user-controlled server (pointed to by the provided URL). Customize/stylize the template document and send directly to the desired target(s). This will trigger a remote reference to the payload document, ultimately pulling and executing the DDE command string.

2. Customize/Stylize the payload document and send directly to the desired target(s).

## References
https://sensepost.com/blog/2017/macro-less-code-exec-in-msword/  
https://staaldraad.github.io/2017/10/23/msword-field-codes/

## Additional Thanks
[@ryHanson](https://twitter.com/ryhanson)  
[@SecuritySift](https://twitter.com/securitysift)  
[@GossiTheDog](https://twitter.com/gossithedog)