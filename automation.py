# using python 3.9

'''
author: Aliasgher Nooruddin
email: aliasghernooruddin@gmail.com
github: https://github.com/aliasghernooruddin
'''

from redmail import outlook
from pathlib import Path
import pandas as pd
import os


''' 
    searches for the PDF files in the current folder and adds them as attachments, you can remove this condition
    if you want to include all the file types or extend this condition depending on your use case 
'''
def getAttachments():
    files = os.listdir()
    attachments = {}

    for i in files:
        if i.endswith('pdf'):
            attachments[i] = Path(i)

    return attachments


def sendEmail(to, cc, html, attachments):
    try:
        outlook.send(
            receivers=to,
            cc=cc,
            subject="MATERIAL DELIVERY DETAILS",
            html=html,
            attachments=attachments
        )

        print("Emails have been sent successfully!!!")

    except Exception as e:
        print(e)


def getRecepients(df, attachments):
    for row in df.iterrows():
        to = row[1]['To'].split(",")
        cc = row[1]['CC'].split(",")
        date = row[1]['Date']
        date = str(date.strftime("%d %B %Y"))

        '''
        opens the HTML format which will be used to send emails
        '''
        with open('index.html') as f:
            lines = f.read()

        html = lines.split("{date}")
        html = html[0] + str(date) + html[1]

        sendEmail(to, cc, html, attachments)


def startScript():
    attachments = getAttachments()

    # gets all the receipients, CC emails from the excel file to whom emails will be sent
    df = pd.read_excel('recepients.xlsx')

    # initializes outlook credentials
    outlook.username = "test@example.com"  # Enter your email
    outlook.password = "****************"  # Enter your password

    getRecepients(df, attachments)


if __name__ == '__main__':
    startScript()
