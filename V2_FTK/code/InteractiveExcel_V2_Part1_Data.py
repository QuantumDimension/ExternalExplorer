########################################################################################
### The Interactive Excel: External Explorer Script Part 1: Data Preparation         ###
### Creator: Glenn Nor                                                               ###
### Date of publication: January 2022                                                ###
### Version 2.13                                                                     ###
### This version is designed to be used with data from the digital forensic software ###
### AccessData Forensic ToolKit 7.5                                                  ###
########################################################################################

import extract_msg, glob, os

###########################################################################
# Script Functions -  Helpers START                                       #
###########################################################################

def Generate_Header(TITLE):
    Head = '''<!DOCTYPE html>
    <HTML>
    <HEAD>
    <META http-equiv="Content-Type" content="text/html"; charset="UTF-8">
    <TITLE>{0}</TITLE>'''.format(TITLE)
    
    HeaderEnd = '''<STYLE type="text/css">
    BODY { font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 10pt; }
    PRE { margin: 0; }
    TABLE.ad { padding: 1px; background-color: #C0C0C0; }
    TABLE.ad TR.top { background-color: #B0D0F0; }
    TABLE.ad TR.odd { background-color: #F8F8F8; }
    TABLE.ad TR.even { background-color: #F0F0F0; }
    TABLE.ad TD, TABLE.ad TH { padding: 3px; vertical-align: top; }
    TH.adlbl, TH.adlbl10, TH.adlbl20 { text-align: left; font-weight: bold; }
    TH.adlbl10 { width: 10%; }
    TH.adlbl20 { width: 20%; }
    .emlhdr, .emlbdy, .emlftr { width: 100%; padding: 0px; background-color: white; }
    .emlbdy { background-color: transparent; }
    .emlhdr TD, .emlhdr TH, .emlbdy TD, .emlftr TH, .emlftr TD { padding: 1px; vertical-align: top; }
    .emlhdr TH, .emlftr TH { padding-right: 100px; }
    .emlhdr, .emlftr { border: none; padding-top: 12px; margin-bottom: 34px; font-size: 10pt; font-family: Arial; }
    .emlhdr { border-top: solid black 4px; }
    .emlftr { border-top: dashed black 1px; margin-top: 12px; }
    .emlftr TR TD PRE { font-family: Verdana, Arial, Helvetica, sans-serif; }
    .emlbdy TR TD PRE { font-family: Verdana, Arial, Helvetica, sans-serif; white-space: pre-wrap; word-wrap: break-word; }
    </STYLE>
    </HEAD>\n'''
    HeaderFull = Head+HeaderEnd
    return HeaderFull

def Generate_HTML_Body_Preamble():
    return '''<BODY>\n<TABLE cellspacing=0 class="emlhdr"><TBODY>\n'''

def Generate_HTML_FROM(MSG_FROM):
    HTML_FROM = '''<TR class="even"><TH class="adlbl10"><NOBR>From: </NOBR></TH><TD>{0}</TD></TR>\n'''.format(MSG_FROM)
    return HTML_FROM

def Generate_HTML_TO(MSG_TO):
    HTML_TO = '''<TR class="even"><TH class="adlbl10"><NOBR>To: </NOBR></TH><TD>{0}</TD></TR>\n'''.format(MSG_TO)
    return HTML_TO

def Generate_HTML_DATE(MSG_DATE):
    HTML_DATE = '''<TR class="odd"><TH class="adlbl10"><NOBR>Sent: </NOBR></TH><TD>{0}</TD></TR>\n'''.format(MSG_DATE)
    return HTML_DATE

def Generate_HTML_SUBJECT(MSG_SUBJECT):
    HTML_SUBJECT = '''<TR class="odd"><TH class="adlbl10"><NOBR>Subject: </NOBR></TH><TD>{0}</TD></TR>\n'''.format(MSG_SUBJECT)
    return HTML_SUBJECT

def Generate_HTML_BODY_POST():
    return '''</TBODY></TABLE>\n'''

def Generate_HTML_BODY_MAINPREAMBLE():
    return '''<TABLE cellspacing=0 class="emlbdy"><TBODY>\n<TR class="even"><TD>\n<div class="WordSection1">\n'''

def HTML_ASSIST_ATTACHMENTS(ATTACH_FILE_NAME, ORIGINAL_FILE_NAME):
    return """<a href="./%s">%s</a><br>""" % (ATTACH_FILE_NAME, ORIGINAL_FILE_NAME)

def Generate_HTML_ATTACHMENTS(count, emailsource):
    Start = """<TD>"""
    Middle = ""
    
    for g in range(count):
        Attachment_Visible_Name = emailsource.attachments[g].longFilename
        Attachment_Hyperlink_Name = emailsource.filename.lower().split(".msg")[0].capitalize().split("\\")[-1]
        Attachment_Extension = Attachment_Visible_Name.split(".")[-1]
        HyperLink = "%s.%s.%s" % (Attachment_Hyperlink_Name, g+1, Attachment_Extension)
        Middle+=HTML_ASSIST_ATTACHMENTS(HyperLink, Attachment_Visible_Name)

        # Save the attachments with correct new filenames
        AttachSave = open("dst\\"+HyperLink, "wb")
        AttachData = emailsource.attachments[g].data
        AttachSave.write(AttachData)
        AttachSave.close()
        
    Attach_HTML = """<TR class="even"><TH class="adlbl10"><NOBR>Attachments: </NOBR></TH><TD>%s</TR>\n""" % Middle
    return Attach_HTML

def ExternalScript():
    # Genereate HTML to PDF script lines
    Source = glob.glob("src\\*.*")
    HTML_Targets = glob.glob("dst\\*.html")
    MSG_Remove_Targets = glob.glob("src/*.msg")
    Regular_Targets = [x for x in Source if x not in MSG_Remove_Targets]
    converter_path = "C:\\Program Files\\wkhtmltopdf\\bin"
    External = open("Temp_Convert_And_Cleanup.bat", "w")

    for HTML_File in HTML_Targets:
        ConvertLine = "%sC:\\Program Files\\wkhtmltopdf\\bin\\wkhtmltopdf.exe%s --keep-relative-links %s %s.pdf\n" % (chr(34), chr(34), HTML_File, HTML_File.split(".")[0])
        External.write(ConvertLine)

    for HTML_File_Delete in HTML_Targets:
        DeleteLine = "del /F {0}{1}{2}\n".format(chr(34), HTML_File_Delete, chr(34))
        External.write(DeleteLine)
    
    for msg_temp in MSG_Remove_Targets:
        delete = "del /F {0}{1}{2}\n".format(chr(34), msg_temp, chr(34))
        External.write(delete)

    for regular in Regular_Targets:
        move = "move {0}{1}{2} {3}{4}{5}\n".format(chr(34), regular, chr(34), chr(34), "dst\\"+regular.split("\\")[-1], chr(34))
        External.write(move)

    InteractiveSpace = "mkdir InteractiveExcel\\items\n"
    MoveAllToNewSpace = "move dst\\*.* InteractiveExcel\\items\\\n"

    External.write(InteractiveSpace)
    External.write(MoveAllToNewSpace)
    
    External.close()

# Script Functions - Helpers END ############################################


###########################################################################
# Script Functions - Main START                                           #
###########################################################################

def Generate_HTML_Email(MSG_SOURCE):
    
    # Email Elements
    Email = extract_msg.Message(MSG_SOURCE)
    Email_ID = Email.filename.lower().split(".msg")[0].capitalize().split("\\")[-1]
    Title = Email_ID.split(".")[0]
    
    try:
        To = Email.to.replace("<", "[").replace(">", "]")
    except:
        To = Email.to
        
    try: 
        From = Email.sender.replace("<", "[").replace(">", "]")
    except:
        From = Email.sender
    
    Date = Email.date
    Subject = Email.subject
    Message = Email.htmlBody

    # HTML Elements
    HTML_Header = Generate_Header(Title)
    HTML_Body_Pre = Generate_HTML_Body_Preamble()
    HTML_From = Generate_HTML_FROM(From)
    HTML_To = Generate_HTML_TO(To)
    HTML_Date = Generate_HTML_DATE(Date)
    HTML_Subject = Generate_HTML_SUBJECT(Subject)
    HTML_BodyPost = Generate_HTML_BODY_POST()
    HTML_BodyMainPre = Generate_HTML_BODY_MAINPREAMBLE()

    # Some emails will not decode with UTF 8. We use a backup just in case.
    try:
        HTML_Message = Message.decode("utf-8")
    except:
        HTML_Message = Message.decode("cp1250")
    
    HTML_Message_End = '''</div>'''

    Email_Attachment_Exists = "FALSE"

    # Check for attachments. Generate Attachment HTML if exist.
    Email_Attachment_Count = len(Email.attachments)
    if (Email_Attachment_Count > 0):
        HTML_Attachments = Generate_HTML_ATTACHMENTS(Email_Attachment_Count, Email)
        Email_Attachment_Exists = "TRUE"

    # HTML Construction
    if (Email_Attachment_Exists == "FALSE"):
        Full_HTML_Email = HTML_Header+HTML_Body_Pre+HTML_From+HTML_To+HTML_Date+HTML_Subject+HTML_BodyPost+HTML_BodyMainPre+HTML_Message+HTML_Message_End
    if (Email_Attachment_Exists == "TRUE"):
        Full_HTML_Email = HTML_Header+HTML_Body_Pre+HTML_From+HTML_To+HTML_Date+HTML_Subject+HTML_Attachments+HTML_BodyPost+HTML_BodyMainPre+HTML_Message+HTML_Message_End
    
    
    HTML_FILENAME = "{0}.html".format(Email_ID)
    
    Generate_File = open("dst\\"+HTML_FILENAME, "w", encoding="utf-8")
    Generate_File.write(Full_HTML_Email)
    Generate_File.close()
    
# Script Functions - Main END ###############################################

###########################################################################
# MAIN PROGRAM                                                            #
###########################################################################

Email_Targets = glob.glob("src/*.msg")

for v in range(len(Email_Targets)):
    Generate_HTML_Email(Email_Targets[v])
    print("Currently Rebuilding Email {0} of {1}".format(v+1, len(Email_Targets)))

print("\nGeneration of HTML Emails with interactive attachment links complete.")

# Create External Script to perform Convertions and Cleanup
print("\nGenerating External Script for creating PDF files and cleaning up.")
ExternalScript()

print("Done.")
