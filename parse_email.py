# pip install win32com
# pip install os
# pip install re

import win32com.client
from os import walk
import re

#
# Get names of all messages in the folder
#

folderpath = "C:\\...\\"

f = []
for (dirpath, dirnames, filenames) in walk(folderpath):
    f.extend(filenames)

#print(f)

#
# Open + Parse emails
#

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

for email in f:
    #print(email)

    try:

        msg = outlook.OpenSharedItem(folderpath + email) 

        #print(msg.SenderName)
        #print(msg.SenderEmailAddress)
        #print(msg.SentOn)
        #print(msg.To)
        #print(msg.CC)
        #print(msg.BCC)
        #print(msg.Subject)
        print(msg.Body)
        #print(msg.Categories)

        del msg

    except:
        print("can't open email")

del outlook