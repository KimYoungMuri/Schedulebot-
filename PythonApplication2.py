from __future__ import print_function
from googleapiclient.discovery import build
from httplib2 import Http
from oauth2client import file, client, tools
import smtplib
from email.mime.text import MIMEText
from time import gmtime, strftime, localtime, sleep

# If modifying these scopes, delete the file token.json.
SCOPES = 'https://www.googleapis.com/auth/spreadsheets.readonly'

# The ID and range of a sample spreadsheet.
SPREADSHEET_ID = '1xtfkX5h9vFIx2XwVvVxIIENpBgKUapkavCO6h99aCKE'
RANGE_NAME = 'Schedule!A2:B'


#emails = ["kpark21@student.kis.or.kr"]
mail_user = "schedulebot@student.kis.or.kr"
password = open("password.txt", "r")
password = password.read()
password = "distcodegwlm"
print("started")
isSent = False
msg = "test"
student = 0
formResponses = 'data.csv!A2:M'
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
        rec_email = email[0]
        weekdayName = strftime("%A",localtime())
        #print(weekdayName)
        #print("Sending soon")
        if email[11] == "Yes":
            #print(rec_email)
            if email[12] == "High School":
                if weekdayName == "Sunday" or "Tuesday" or "Thursday":
                    second = "High School Lunch"
                    first = "club"
                  
                elif weekdayName =="Monday":
                    second = "High School Lunch"
                    first = "Advisory"
                elif weekdayName == "Wednesday":
                    second = "Contact Time"
                    first = "High School Lunch"
            elif email[12] == "Middle School":
                first = "Middle School Activity"
                second = "Middleschool Lunch"
            if dayName == "a":
                sbj = email[3]+", "+email[4]+", "+email[5]+", "+email[6]
                msg = "8:00 - 9:21:\t" + email[3]+"\n" + "9:28 - 10:48:\t" + email[4]+"\n" + "10:50 - 11:25:\t" + first+ "\n" + "11:30 - 12:08:\t" + second+"\n" + "12:13 - 1:33:\t" + email[5]+"\n" + "1:40 - 3:00:\t" + email[6]
            elif dayName == "b":
                sbj = email[7]+", "+email[8]+", "+email[9]+", "+email[10]
                msg = "8:00 - 9:21:\t" + email[7]+"\n" + "9:28 - 10:48:\t" + email[8]+"\n" + "10:50 - 11:25:\t" + first+ "\n" + "11:30 - 12:08:\t" + second+"\n" + "12:13 - 1:33:\t" + email[9]+"\n" + "1:40 - 3:00:\t" + email[10]
            elif dayName == "c":
                sbj = email[6]+", "+email[5]+", "+email[4]+", "+email[3]
                msg = "8:00 - 9:21:\t" + email[6]+"\n" + "9:28 - 10:48:\t" + email[5]+"\n" + "10:50 - 11:25:\t" + first+ "\n" + "11:30 - 12:08:\t" + second+"\n" + "12:13 - 1:33:\t" + email[4]+"\n" + "1:40 - 3:00:\t" + email[3]
            elif dayName == "d":
                sbj = email[10]+", "+email[9]+", "+email[8]+", "+email[7]
                msg = "8:00 - 9:21:\t" + email[10]+"\n" + "9:28 - 10:48:\t" + email[9]+"\n" + "10:50 - 11:25:\t" + first+ "\n" + "11:30 - 12:08:\t" + second+"\n" + "12:13 - 1:33:\t" + email[8]+"\n" + "1:40 - 3:00:\t" + email[7]
            elif dayName == "e":
                sbj = email[4]+", "+email[3]+", "+email[6]+", "+email[5]
                msg = "8:00 - 9:21:\t" + email[4]+"\n" + "9:28 - 10:48:\t" + email[3]+"\n" + "10:50 - 11:25:\t" + first+ "\n" + "11:30 - 12:08:\t" + second+"\n" + "12:13 - 1:33:\t" + email[6]+"\n" + "1:40 - 3:00:\t" + email[5]
            elif dayName == "f":
                sbj = email[8]+", "+email[7]+", "+email[10]+", "+email[9]
                msg = "8:00 - 9:21:\t" + email[8]+"\n" + "9:28 - 10:48:\t" + email[7]+"\n" + "10:50 - 11:25:\t" + first+ "\n" + "11:30 - 12:08:\t" + second+"\n" + "12:13 - 1:33:\t" + email[10]+"\n" + "1:40 - 3:00:\t" + email[9]
            elif dayName == "Final":
                sbj = "Last Schedule Bot email :( Algebra II"
                msg = "Oh wow this is emotional, Thank you for supporting my project. It's been one hell of a ride :')."+"\n"+" If you want to sign up for schedule bot again, I'll make an updated version for next year."+"\n"+" If you have any feedback just send it to me through this email."+"\n"+"Anyways for those of you who have finals,"+"\n\n"+"8:00 - 9:30: Algebra 2"+"\n"+"Good Luck! Best Wishes from Schedule Bot!"+"\n"+"For Exam Locations Please Check Schoology or https://bit.ly/2WUZE3q"
            elif dayName == "PepRally6587":
                if email[12] == "High School":
                    sbj = email[8]+", "+email[7]+", "+email[10]+", "+email[9]+" Pep Rally"
                    msg = "8:00 - 9:05: " + email[8]+"\n" + "9:12 - 10:17: " + email[7]+"\n" + "10:24 - 11:02: " + first+ "\n" + "11:09 - 11:44: " + second+"\n" + "11:51 - 12:58: " + email[10]+"\n" + "1:05 - 2:13: " + email[9] + "\n2:20 - 3:00: Winter Pep Rally"
                elif email[12] == "Middle School":
                    sbj = email[8]+", "+email[7]+", "+email[10]+", "+email[9]
                msg = "8:00 - 9:21: " + email[8]+"\n" + "9:28 - 10:48: " + email[7]+"\n" + "10:53 - 11:25: " + first+ "\n" + "11:30 - 12:08: " + second+"\n" + "12:13 - 1:33: " + email[10]+"\n" + "1:40 - 3:00: " + email[9]
            elif dayName == "Special12":
                sbj = email[3] +", "+email[4]
                msg = "8:00 - 9:21: " + email[3]+"\n" + "9:28 - 10:08: " + first+"\n" + "10:08 - 10:48: School Photo" + "\n" + "10:55 - 12:15: " + email[4]+"\n" + "Parent Teacher Conferences"
            elif dayName == "Special34":
                sbj = email[5] +", "+email[6]
                msg = "8:00 - 9:21: " + email[5]+"\n" + "9:28 - 10:08: " + first+"\n" + "10:08 - 10:48: School Photo" + "\n" + "10:55 - 12:15: " + email[6]+"\n" + "Parent Teacher Conferences"
            elif dayName == "aprilFools":
                sbj = "School is not cancelled, April Fools!"
                msg = "April fools! While you are here, click this link to fill out my schedule bot survey!\nwww.everyview.net/school-cancelled/"
            elif dayName == "EE":
                sbj = "EE trip tomorrow!"
                msg = "Have fun at your EE trip!"
            elif dayName == "a Half":
                sbj = "It's Half-loween! Dress up tomorrow!, "+email[3]+", "+email[4]+", "+email[5]+", "+email[6]
                msg = "8:00 - 8:50: " + email[3] +"\n"+ "8:55 - 9:45: " + email[4] + "\n"+ "9:50 - 10:40: " + email[5] + "or MS Lunch" +"\n"+ "10:40 - 11:20: " + email[5] + "or HS Lunch" +"\n"+ "11:25 - 12:15: " + email[6] +"\n"+"Parent teacher conferences :| oof"
            #sbj = "Update your schedules with your 2nd semester courses!"
            #sbj = "Message about the Corona Virus from the School Director"
            

            if dayName != 'x':
                
                server = smtplib.SMTP("smtp.googlemail.com:587")
                server.ehlo()   
                server.starttls()
                server.login(mail_user, password)
                text = MIMEText("Hello "+email[2]+",\n"+"Here's your virtual schedule for tomorrow \n\n"+msg+"\n\nCurrently "+str(len(values))+" people at home are using Virtual Bot"+"\n\nThank you for using Virtual Bot\nCreated by Kevin Park and Paul Ericksen.\n\nIf you wish to unsubscribe, fill out No for the last check box.\nhttps://www.everyview.net/schedule-bot-form.html")
                
                text['Subject'] = sbj
                text['To'] = rec_email
                text['From'] = mail_user
                server.sendmail(mail_user,rec_email,text.as_string())
                #rec_email
                server.quit()
                sleep(3)
                #break
                #print("sent")
                


                
           
       
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
#"""
while True:
    currentTime = int(strftime("%H",localtime()))
    
    if currentTime == 18:#change to 17 after finals
        if isSent == False:
            isSent = True
            date = strftime("%j",localtime())
            date = date.lstrip("0")
            whatDayIsIt()
            #sleep(84600) #Sleep for 23 1/2 hours
            #print(currentTime)
    elif currentTime == 19:
        isSent = False
"""
        
    
if __name__ == '__main__':
    date = strftime("%j",localtime())

    date = date.lstrip("0")
    print(date)
    whatDayIsIt()
"""

