# ExternalExplorer
The Interactive Excel: External Explorer - A python script for creating an Excel "eDiscovery" platform (sort of).
___________

<b>About Version 1:</b> This is designed to be used with data exported from Forensic Toolkit 7.5 using the FTK Plus / Qview function. The exported data needs to be put in specific folders, but after that the script can be run, and the rest is automated.

[1] Export the files you want via the options "Create portable case".</br>
[2] Go the case folder and copy "files" and "data.db" to the "src" folder in the script structure.</br>
[3] Export a file list from FTK (only showing the files you want in the file viewer) directly to the "src" folder in the script structure.</br>
[4] Run the batch launcher and let the automation take over.</br>

The finished product is this:

1. InteractiveExcel [Folder]</br>
1.1. Interactive_Excel_Prototype.xlsx</br>
1.2. items [Folder]</br>
--->1.2.1. [all the files created from part 1 of the script]</br>

The excel file will have metadata, one line for each file, showing for example, email from, email to etc. if its an email, and file size, extension etc. if its a file.

In column B, there will be hyperlinks, which allows you to open each file directly from Excel.

The PDF versions of the emails will also have the ability to open attachments directly from the file.
_____________
<b>About Version 2:</b> This is designed to be used with data exported from Forensic Toolkit 7.5 using the normal export function. Emails need to be exported as MSG (there is an option in the export settings).

[1] Export the files you want. Check the option to have Emails as MSG, and Item Numbers as file names</br>
[2] Move the files to the "src" folder in the script structure.</br>
[3] Export a file list from FTK (only showing the files you want in the file viewer) directly to the main folder in the script structure.</br>
[4] Run the batch launcher and let the automation take over.</br>

The finished product is this:

1. InteractiveExcel [Folder]</br>
1.1. Interactive_Excel_Prototype.xlsx</br>
1.2. items [Folder]</br>
--->1.2.1. [all the files created from part 1 of the script]</br>

The excel file will have metadata, one line for each file, showing for example, email from, email to etc. if its email, and file size, extension etc. if its a file.

In column B, there will be hyperlinks, which allows you to open each file directly from Excel.

The PDF versions of the emails will also have the ability to open attachments directly from the file.
_______________

<h2>A detailed article explaining how to use this script can be found here:</h2>
<h3>www.linkedin.com</h3>
