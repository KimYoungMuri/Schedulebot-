
from __future__ import print_function
from googleapiclient.discovery import build
from httplib2 import Http
from oauth2client import file, client, tools
import smtplib
from email.mime.text import MIMEText
from time import gmtime, strftime, localtime, sleep
import random

# If modifying these scopes, delete the file token.json.
SCOPES = 'https://www.googleapis.com/auth/spreadsheets.readonly'

# The ID and range of a sample spreadsheet.
SPREADSHEET_ID = '1xtfkX5h9vFIx2XwVvVxIIENpBgKUapkavCO6h99aCKE'
RANGE_NAME = 'Schedule!A2:B'


#emails = ["kpark21@student.kis.or.kr"]
mail_user = "schedulebot@student.kis.or.kr"
password = "distcodegwlm"
print("started")
isSent = True
msg = "test"
student = 0
formResponses = 'tempProj!A2:C'
def main(dayName):
    """Shows basic usage of the Sheets API.
    Prints values from a sample spreadsheet.
    """
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    store = file.Storage('token.json')
    creds = store.get()
    if not creds or creds.invalid:
        flow = client.flow_from_clientsecrets('credentials.json', SCOPES)
        creds = tools.run_flow(flow, store)
    service = build('sheets', 'v4', http=creds.authorize(Http()))

    # Call the Sheets API
    sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId=SPREADSHEET_ID, range=formResponses).execute()
    values = result.get('values',[])
    
    for email in values:
        ranNum = random.randint(100000,999999)
        rec_email = email[1]
        ranNum = email[2]
        server = smtplib.SMTP("smtp.googlemail.com:587")
        server.ehlo()   
        server.starttls()
        server.login(mail_user, password)
        text = MIMEText("Dear "+email[0]+",\n\n"+"Thank you for your interest in participating in this survey. Your participation is greatly appreciated." +"\n\n"+"If you are under the age of 18, please have your parents sign this(https://docs.google.com/document/d/1bRGbqTWUNcw_ZfQ2zGk-oevF01M2YZaqFquieDL-Iu0/edit) consent letter and email schoe21@student.kis.or.kr before completing the survey. There is more important information in the form, so please read the form even if you do not have to fill one out. Completion of this form means you are agreeing to having your data be used in this study, which includes the survey data and your GPA. However, the researcher will not know your GPA. This form is not required if you are 18 or above (Korean or American Age)."+"\n\n"+"Your ID Number is ("+str(ranNum)+"). Please complete the form with this ID Number. " +"\n\n"+"Please complete the form and survey as soon as possible."+"\n\n"+"Thank you."+"\n\n\n"+"Powered by an unpaid Kevin Park")
        text['Subject'] = "Survey"
        text['To'] = rec_email
        text['From'] = mail_user
        server.sendmail(mail_user,rec_email,text.as_string())
        
        server.quit()
        sleep(1)

        #print("sent")


                
           
"""     
def whatDayIsIt():
    
    store = file.Storage('token.json')
    creds = store.get()
    if not creds or creds.invalid:
        flow = client.flow_from_clientsecrets('credentials.json', SCOPES)
        creds = tools.run_flow(flow, store)
    service = build('sheets', 'v4', http=creds.authorize(Http()))
    sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId=SPREADSHEET_ID, range=RANGE_NAME).execute()
    values = result.get('values',[])
    for day in values:

        if day[0] == date:
            main(day[1])
            #print (day[1])
            
        else:
            pass
"""
    
if __name__ == '__main__':
    date = strftime("%j",localtime())

    date = date.lstrip("0")
    print(date)
    main("a")
    #whatDayIsIt()
