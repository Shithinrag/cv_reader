import os
import docx2txt
import re
from win32com import client as wc

os.system('cls')
script_path = os.path.dirname(__file__)+"/"



print(" Python3 Script to extract Name, Email, Address, Phone Number from docx ")
print(" doc file will be converted to docx ")
print()
print()
print()

file_name = input("Enter file name with extension: ")
print(script_path + file_name)
if os.path.isfile( script_path + file_name ):
    print("File Exists")
else:
    print(f"\'{file_name}\' Doesn't exist. Script Ended")
    exit()
    
if print(file_name[-4:]) is not "docx":
    w = wc.Dispatch('Word.Application')
    w.visible = False
    doc=w.Documents.Open( script_path + file_name )

    doc.SaveAs(script_path + file_name[:-3] + "docx",16)
    file_name = file_name[:-3] + "docx"


document = docx2txt.process(script_path + file_name)
document = document.split()
# print(document)

name = ""
email = ""
address = ""
phone = ""

# name = document[0] + " " + document[1]

addressFlag = False
nameFlag = False

for i in range(10):
    if re.search("name", document[i].lower()) is not None:
        nameFlag = True
        continue
    elif document[2] == "Page":
        name = document[3] + " " + document[4]
    if nameFlag:
        if document[i].isalpha():
            name += document[i]
            break



for text in document:
    if phone == "":
        rePhone = re.search(r'[0-9]{10}',text)
        if rePhone is not None:
            phone =  text[rePhone.start() : rePhone.end()]
    if email == "":
        reEmail = re.search(r"[a-z0-9\.\-+_]+@[a-z0-9\.\-+_]+\.[a-z]+",text)
        if reEmail is not None:
            email =  text[reEmail.start() : reEmail.end()]
    
    if re.search("Address",text) is not None:
        addressFlag = True
    
    if addressFlag:
        address += text + " "
        if re.search("[0-9]{6}", text) is not None or re.search("india", text.lower()) is not None:
            addressFlag = False
        
os.system('cls')
print("Name: " + name)
print("Email: " + email)
print("Address: " + address)
print("Phone: " + phone)