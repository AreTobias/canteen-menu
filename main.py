import requests as rq
from bs4 import BeautifulSoup
import win32com.client


# The url for the catering's website

URL = "https://fristeren.no/"

page = rq.get(URL)


soup = BeautifulSoup(page.content, "html.parser")

# The information about the menu is under ID 'Section2'

result = soup.find(id = "Section2")

# None of the 'p' tags had a name, so selecting all of them

uke = result.findAll("p")

# Adds every 'p' tag to an array to be accessible with indexes 
text = []
for i in uke:
    text.append(i.text) 


# Meny array holding the information about the menu
meny = []
meny.append(text[5])
meny.append(text[6])
meny.append(text[7])
meny.append(text[8])
meny.append(text[9])
meny.append(text[10])
meny.append(text[11])

# Getting the week number based on the website
uke = ""
uke += text[5][-3]
uke += text[5][-2]




'''

This section is taking the parsed data and sending it as an email
to the selected recipients using win32com module

'''


# Creating an object holding the Outlook application
ol = win32com.client.Dispatch('Outlook.Application')

# Setting the memory size for the email
olmailitem = 0x0

# Creating the e-mail object
newmail = ol.CreateItem(olmailitem)

# Setting the email subject
newmail.Subject = f"Kantine meny uke {uke}"


# Creating an empty string to hold the e-mail body
menyStr = ""

# adding the body information to the string
for i in meny:
    menyStr += i + "\n\n"

menyStr += "\nDenne e-posten er auto generert"

# Adding the string to the e-mail body object
newmail.Body = menyStr

employees = []

with open("ansatte.txt", "r") as rf:
    lines = rf.readlines()

    for i in lines:
        i.strip("\n")
        employees.append(i)


# Adding the members from employee array to recipients
newmail.To = ";".join(employees)

# This gives a preview of the e-mail before its sent.
# Including recipients, subject and the body.
newmail.Display()

# The last call is to send the e-mail
#newmail.Send()




