# Outlook .msg File Python Parser

<img width="85" alt="map-user" src="https://img.shields.io/badge/views-1657-green"> <img width="125" alt="map-user" src="https://img.shields.io/badge/unique visits-781-green">

Email(s) saved with the .msg file extension are not human read-able when opened in a text editor. These files (.msg) are also difficult to process in Python using common text manipulation process ex. regular expressions. This repository has a python script that can parse all .msg files in a folder and can extract the following fields

- SenderName
- SenderEmailAddress
- SentOn
- To
- CC
- BCC
- Subject
- Body
- Categories

The python script has several dependencies. The python library is required. [win32com](https://pypi.org/project/pywin32/). The [win32com](https://pypi.org/project/pywin32/) library is used to access ```Outlook.Application``` requiring you to have the outlook email client installed on the **windows** machine you run the Python script from.

To run the python script

1. ```pip install win32com```
2. Set the ```folderpath``` variable in the python script to the path of a folder with the .msg file you want to process
3. Comment / uncomment the print statements you want / don't want depending on which properties of the email you want to show
