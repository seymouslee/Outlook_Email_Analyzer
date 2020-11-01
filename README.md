# Email analyzer macro for Outlook

This macro checks if the attachments in the email are potentially malicious. In malware phishing attacks, attackers would masqurade malware file extension such executables (.exe) and Java Archive file (.jar) as another file extension such as PDF or text files to appear normal to the victim. Our macro checks the attachment's actual file extension by checking its certificate. If the file turns out to be malicious, our macro will warn the user and save the file certificate and a report of the email for future investigation.

You can view our project demostration here: https://www.youtube.com/watch?v=NncG8rnJCPA&feature=youtu.be

## How to add the macro in Microsoft Outlook

1. Press Alt + F11 to open up the Visual Basic window
2. Right Click on "Project1" or any other project in the Project explorer panel on the left
   - Select "Import File".
   - Then select the file "Outlook_Email_Analyzer.bas" from the repository.
3. Save and close the Visual Basic window.
4. On the home page, click on File > options > Customise Ribbon
   - On the right panel "Main Tabs", right click on Home and select "Add New Group"
   - On the left side, select "Macros" in the combo/ selection box
   - Select "<projectname>.Extract" and click on the "Add" button in the middle
     - make sure the New group that you have just added is highlighted.
   - you can choose to rename the macro like "Check Attachment" for easier reference
5. when you go back to the home page, there should be the macro displayed on the ribbon under the "Home" tab!
6. Select an email you would like to check and run the macro.
   
## Notes:

- if you would like to re-create attachments shown in the video, the files are stored in the folder called "FileDemo".
- if there are any errors with the file path, please open the "Document" directory in file explorer and try again.

- Team Inconspicuous 
