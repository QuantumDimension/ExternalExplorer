#######################################################################################
### The Interactive Excel: External Explorer Script Part 1: Data Preparation        ###
### Creator: Glenn Nor                                                              ###
### Date of publication: January 2022                                               ###
### Version 1.15                                                                    ###
### This version is designed to be used with data from FTK Plus (formerly Qview)    ###
### A software that is part of AccessData Forensic ToolKit 7.5 Software.            ###
#######################################################################################


### IMPORT MODULES ###
import shutil
import sqlite3
import datetime
import pytz
import os
import glob
######################


### Change current directory to SRC and create items folder ###
os.chdir("src")
os.mkdir("items")
###############################################################

### Functions to replace attachment HTML section with version containing active hyperlink ##########################
def HTML_ASSIST_ATTACHMENTS(ATTACH_FILE_NAME, ORIGINAL_FILE_NAME):
    return """<a href="./%s">%s</a><br>""" % (ATTACH_FILE_NAME, ORIGINAL_FILE_NAME)

def HTML_ATTACHMENTS():
    Start = """<TD>"""
    Middle = ""
    for g in range(len(Attach)):
        HyperLink = "%s.%s.%s" % (Attach[g][3], g+1, Attach[g][1])
        VisibleName = Attach[g][0]
        Middle+=HTML_ASSIST_ATTACHMENTS(HyperLink, VisibleName)
    Attach_HTML = """<TR class="even"><TH class="adlbl10"><NOBR>Attachments: </NOBR></TH><TD>%s</TR>\n""" % Middle

    
    return Attach_HTML
####################################################################################################################

### Connect and fetch relevant data from Qview Portable Database #################################
conn = sqlite3.connect("data.db")
cur = conn.cursor()
cur.execute("SELECT ObjectName,ObjectFiles_Extension,ObjectID,ParentID,filepath,ObjectFiles_FileCategory from objectdata")
data = cur.fetchall()
##################################################################################################

### Overview of Indices ###
# 0 - ObjectName          #
# 1 - Extension           #
# 2 - ObjectID            #
# 3 - ParentID            #
# 4 - SourceLocation      #
# 5 - FileCategory        #
###########################


### All database variables as lists (For future use) ###
Name = [data[i][0] for i in range(len(data))] 
Ext = [data[i][1] for i in range(len(data))]
OID = [data[i][2] for i in range(len(data))]
PID = [data[i][3] for i in range(len(data))]
SRC = [data[i][4] for i in range(len(data))]
CAT = [data[i][5] for i in range(len(data))]
########################################################

### Empty lists for use later ###
Emails = []
Exclusion = []
#################################

### Generate Email list and Exclusion list to help transfer later and separate emails from attachments and files ###
for g in range(len(data)):
    if (data[g][1] == "" or data[g][1] == "MSG"):
        Emails.append(data[g])
        Exclusion.append(data[g][2])
####################################################################################################################

### Convert and move DAT email files to HTML format and items folder ###
for y in range(len(Emails)):
    Original = Emails[y][4]
    ID_Tag = Emails[y][4].split("/")[-1].split(".")[0]
    New = "items/%s.html" % ID_Tag
    shutil.copyfile(Original, New) 
########################################################################

### INFO ###############################################################################
# The emails coming from FTK Plus (Qview Portable) has a bug with regards to           #
# time zone. All emails from FTK Plus export has universal (GMT+0) timezone            #
# This is of course a problem as someone going through these emails are not going to   #
# adjust for the correct timezone in their heads, while they are reading the emails.   #
#                                                                                      #
# To fix this the section below will open each email in HTML format and:               #
# (1)Identify where the timestamp is in the HTML code.                                 #
# (2)Strip the timestamp section apart into chunks.                                    #
# (3)Put the chunks together again, but in a python recognized time format             #
# (4)Assign Universal time / GMT+0 to it                                               #
# (5)Change the python timestamp timezone to Oslo Norway (GMT+1/GMT+2(Summer))         #
#                                                                                      #
#  This means python will automatically correct summer time between March and October. #
# (6)Place the new adjusted python timestamp object back into the HTML code.           #
# (7)Recreating the HTML email, now with corrected time.                               #
#                                                                                      #
#   !PLEASE BE AWARE THAT THIS PROGRAM WILL ALWAYS ADJUST TO OSLO, NORWAY TIMEZONE!    #
#                                                                                      #
########################################################################################

### Open Emails only, strip time and fix/adjust timezone, put HTML code back together ##################
for t in range(len(Emails)):
    print("Currently Rebuilding Email #{0} of {1}".format(t+1, len(Emails))) 
    FileSrc = Emails[t][4]
    FileID = Emails[t][2]
    FileCat = Emails[t][5]
    try:
        ActiveEmail = open(FileSrc, "r", encoding="UTF-8").readlines()
    except:
        print("Could not open email-file: {0}. File type: {1}. If this is not an email type, then all is good.".format(FileSrc, FileCat))
        ActiveEmail = ""
        
    for v in range(len(ActiveEmail)):
        if "<NOBR>Sent:" in ActiveEmail[v]:
            RawHTML = ActiveEmail[v]
            TimeStr = RawHTML.split("<TD>")[-1].split("</TD")[0]
            Date = TimeStr.split(" ")[0]
            Time = TimeStr.split(" ")[1]
            Year = int(Date.split(".")[2])
            Month = int(Date.split(".")[1])
            Day = int(Date.split(".")[0])
            Hour = int(Time.split(":")[0])
            Min = int(Time.split(":")[1])
            Sec = int(Time.split(":")[2])
            CaseTimeZone = "Europe/Oslo"
            Universal = datetime.datetime(Year, Month, Day, Hour, Min, Sec, 0, pytz.UTC)
            CorrectedDateTime = Universal.astimezone(pytz.timezone(CaseTimeZone))
            CorrectedDateTimeStr = "{:%d.%m.%Y %H:%M:%S %z}".format(CorrectedDateTime)
            Pre = """<TR class="odd"><TH class="adlbl10"><NOBR>Sent: </NOBR></TH><TD>"""
            Post = """</TD></TR>\n"""
            NewObject = Pre+CorrectedDateTimeStr+Post
            ActiveEmail[v] = NewObject
            
        if "Attachment" in ActiveEmail[v]:
            Attach = []
            for w in range(len(PID)):
                if PID[w] == FileID:
                    Attach.append(data[w])
                    Exclusion.append(data[w][2])
            AttachIndex = v

            ActiveEmail[AttachIndex] = HTML_ATTACHMENTS()

            for p in range(len(Attach)):
                OriginalAttach = Attach[p][4]
                New = "items/%s.%s.%s" % (FileID, p+1, Attach[p][1])
                try: 
                    shutil.copyfile(OriginalAttach, New)
                except:
                    print("Could not transfer {0}. This might be due to deleted file status".format(OriginalAttach))

        Rebuild = open("items/%s.html" % FileSrc.split("/")[-1].split(".")[0], "w", encoding="UTF-8")

        for j in range(len(ActiveEmail)):
            Rebuild.write(ActiveEmail[j])
        Rebuild.close()
######################################################################################################

### Start tranfer of files that are not EMAILS and not ATTACHMENTS to emails. Important for exports containing files from PC image ###
for f in range(len(data)):
    Current = data[f][2]
    if Current not in Exclusion:
        print("Orphan Item found and transferred...", Current)
        Origin = data[f][4]
        CurrentCat = data[f][5]
        NewLoc = "items/%s.%s" % (data[f][2], data[f][1])
        try:
            shutil.copyfile(Origin, NewLoc)
        except:
            print("Could not transfer {0}. File type: {1}. If this is non-size type, then all is good.".format(Origin, CurrentCat))

shutil.move("items", "..\dst")
######################################################################################################################################
        

### Identify all HTML files, and put into list. Create BATCH file with ready command lines for using WKHTMLTOPDF to convert HTML To PDF, inc. hyperlinks ####################
os.chdir("..\\dst\items")
target_html = glob.glob("*.html")
os.chdir("..\\..")
converter_path = "C:\\Program Files\\wkhtmltopdf\\bin"
GQ = open("Temp_Auto-Convert-HTML-PDF-and-CLEAN.bat", "w")

for elem in target_html:
    T = "%sC:\\Program Files\\wkhtmltopdf\\bin\\wkhtmltopdf.exe%s --keep-relative-links dst\\items\\%s dst\\items\\%s.pdf\n" % (chr(34), chr(34), elem, elem.split(".")[0])
    GQ.write(T)

for tu in target_html:
    R = "del dst\\items\\%s\n" % tu
    GQ.write(R)
GQ.close()
#############################################################################################################################################################################


print("Done...")
