# Office-DDE-Payloads
Collection of scripts and templates to generate Office documents embedded with the DDE, macro-less command execution technique.

## Overview
Leverages the macro-less DDE code execution technique described 
by [@_staaldraad](https://twitter.com/_staaldraad) and [@0x5A1F](https://twitter.com/Saif_Sherei) (blog post link in References 
section below) to generate two malicious Word documents:

*template-final.docx*
- This is the document sent to the target (e.g., via phishing).
  The webSettings configured to pull DDE from payload-final.docx, 
  which is hosted by a server specified by the user. 

*payload-final.docx*
- Contains user-provided DDE payload/command string. Hosted by
  attacker-controlled server (URL provided by user and baked
  into template-final.docx).

Obfuscation and evasion techniques inspired by @_staaldraad
(blog post link in References setion below). 

## References:
https://sensepost.com/blog/2017/macro-less-code-exec-in-msword/
https://staaldraad.github.io/2017/10/23/msword-field-codes/