# GASmerge
A simple Google Sheets mail merge tool using Google Apps Script.  This tool combines the power of Drive, Sheets and Gmail into one mighty little mail merge tool.

This script has two automated processes: 1) A folder link iterator & 2) The mail merge.

1) The link iterator will extract the all of the file links from a Drive folder.  It allows you to paste in the URL of the folder with all of the docs you want to grab the links of.  

2) The Mail Merge will send out emails with the linked file as a PDF attachment.
---
### Using the script for the first time

When you install this script, a new "Sidebar" tab will be added to the Sheets menu.

![Sidebar Screenshot](screenshots/Sidebar.png#center?raw=true "Sidebar")

In order to access the Sidebar, you will need to authorize the script to access the data.  When you open the Sidebar for the first time, you may see an Authorization prompt like this. 

![Authorization Screenshot](screenshots/Authorization.png?raw=true "Authorization")

Because this is a script that you are installing yourself, instead of buying an officially verified third-party script, you will need to acknowledge this fact by clicking on "Advanced"...

![Not Verified Screenshot](screenshots/notVerified.png?raw=true "Not Verified")

...then scroll down to the bottom and click on "Go to Untitled project (unsafe)".

![Unsafe Screenshot](screenshots/Unsafe.png?raw=true "Unsafe")

