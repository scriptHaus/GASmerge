# GASmerge
A simple Google Sheets mail merge tool using Google Apps Script.  This tool combines the power of Drive, Sheets and Gmail into one mighty little mail merge tool.

This script automates two processes: 1) A folder link iterator & 2) A mail merge.

1) The link iterator will extract the all of the file links from a Drive folder.  It allows you to paste in the URL of a Google Drive folder with all of the docs you want to grab the links from.  

2) After you've added the names and emails to the sheet, the mail merge will send out an email with the linked file as a PDF attachment.
---
### Using the script for the first time

When you install this script, a new "Sidebar" tab will be added to the Sheets menu.

![Sidebar Screenshot](screenshots/Sidebar.png#center?raw=true "Sidebar")

In order to access the Sidebar, you will need to authorize the script to access the data.  When you open the Sidebar for the first time, you may see an Authorization prompt like this. 

![Authorization Screenshot](screenshots/Authorization.png?raw=true "Authorization")

Because this is a script that you are installing yourself (instead of buying an officially verified third-party script) you will need to acknowledge this fact by clicking on "Advanced"...

![Not Verified Screenshot](screenshots/notVerified.png?raw=true "Not Verified")

...then scroll down to the bottom and click on "Go to Untitled project (unsafe)".

![Unsafe Screenshot](screenshots/Unsafe.png?raw=true "Unsafe")

This will bring up another screen asking you to verify that you are allowing the script to access your Drive (for the attachments), Sheets (for the merge data) and Gmail (for sending the emails).  There is no third party info used in the script, but it does warn you that the script could include it (if you modified the script to do so).  To proceed, you must click the "Allow" button.

![Allow Screenshot](screenshots/Allow.png?raw=true "Allow")

Once you have finished that, you're ready to go!

---
### Installation

In order to use the script, you need to add two files to your Sheets workbook.   To begin, click on Script Editor from the Tools tab.

![Script Editor Menu Screenshot](screenshots/scriptEditor.png?raw=true "Script Editor")

This will open up the Google IDE where you will delete the sample JS data in the Code.gs tab...

![Sample JS Data Screenshot](screenshots/sampleJS.png?raw=true "Sample JS function")

...and replace it with the code from GASmerge Code.gs.  Be sure to save!

Next, create a new HTML file under the File tab.

![Create HTML file Screenshot](screenshots/addHTMLfile.png?raw=true "Create HTML file")

Be sure to name the new HTML file "Sidebar".

![Name HTML file Screenshot](screenshots/nameHTML.png?raw=true "Name HTML file")

Again, delete the sample HTML code and replace with the GASmerge Sidebar.html code and save.

![Sample HTML file Screenshot](screenshots/sampleHTML.png?raw=true "Sample HTML file")

Once you have saved both files, you are ready to use GASmerge!

---
### Instructions

To open the Sidebar, you may need to refresh the workbook page on your browser.  Once the Sheets workbook has been refreshed you can click on the Sidebar dropdown from the Sidebar tab.

![Sidebar Screenshot](screenshots/Sidebar.png#center?raw=true "Sidebar")

That will open up the Sidebar.

![Open Sidebar Screenshot](screenshots/openSidebar.png#center?raw=true "Open Sidebar")

To list the PDF files you will be using as attachments, paste the URL of the Drive folder containing the PDFs into the input box in the Sidebar, then push the "Get the Links" button.  The Drive folder should look something like this.

![Drive URL Screenshot](screenshots/driveURL.png#center?raw=true "Drive URL")

The script will create a new sheet in your workbook... 

![New Tab Screenshot](screenshots/newTab.png?raw=true "GASmerge Tab")

...and add hyperlinks for all of the documents it finds in the Drive folder.

Next, add the names and email addresses of the contacts that will receive each of the files listed.  (Be sure to distinguish multiple email addresses by using a comma separator). 

Finally, push the "Send Email" button when ready to send. The script will update each row as it sends the emails.
