from __future__ import print_function
from googleapiclient.discovery import build
from httplib2 import Http
from oauth2client import file, client, tools
import smtplib
from email.mime.text import MIMEText
from time import gmtime, strftime, localtime, sleep

SCOPES = 'https://www.googleapis.com/auth/spreadsheets.readonly'
SPREADSHEET_ID = '1arYFxK_4Fc8mvPVLpZhfpsvcKH0EroDY2ErZdTty9fo'
RANGE_NAME = 'Schedule!A2:B'

mail_user = "schedulebot@student.kis.or.kr"
#password = open("/home/code/Desktop/scheduleBot/password.txt", "r")
#password = password.read()
password = "@Phoenix2021"
print("started")
isSent = False
msg = "test"
student = 0
formResponses = 'data.csv!A2:AH'


def finding_letter_day(date):
    store = file.Storage('/home/code/Desktop/scheduleBot/token.json')
    creds = store.get()
    if not creds or creds.invalid:
        flow = client.flow_from_clientsecrets('/home/code/Desktop/scheduleBot/credentials.json', SCOPES)
        creds = tools.run_flow(flow, store)
    service = build('sheets', 'v4', http=creds.authorize(Http()))
    sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId=SPREADSHEET_ID, range=RANGE_NAME).execute()
    values = result.get('values',[])
    for day in values:
        if day[0] == date:
            # print("HERE!")
            # print (day[1] + "Today's day[1]")
            return(day[1])
        else: pass

def club_or_adv(dayName, email):
    if dayName == "Sunday":
        lst = [email[20], email[21]] #monday club + link
    elif dayName == "Tuesday":
        lst = [email[22], email[23]] #wednesday club + link
    elif dayName == "Thursday":
        lst = [email[24], email[25]] #friday club + link
    elif dayName == "Monday" or dayName == "Wednesday":
        lst = ["Advisory", ""]
    return lst

def block_names(day_type, email):
    if day_type == "a":
        class_names = [email[4], email[6], email[8], email[10]] #1,2,3,4
    elif day_type == "b":
        class_names = [email[12], email[14], email[16], email[18]] #5,6,7,8
    elif day_type == "c":
        class_names = [email[10], email[8], email[6], email[4]] #4,3,2,1
    elif day_type == "d":
        class_names = [email[18], email[16], email[14], email[12]] #8,7,6,5

    #special days
    elif day_type == "Final":
        class_names = ["Good","luck","on","finals!"]
    return class_names

def block_links(day_type, email):
    if day_type == "a":
        class_links = [email[5], email[7], email[9], email[11]]
    elif day_type == "b":
        class_links = [email[13], email[15], email[17], email[19]]
    elif day_type == "c":
        class_links = [email[11], email[9], email[7], email[5]]
    elif day_type == "d":
        class_links = [email[19], email[17], email[15], email[13]]
    return class_links

def lunch_location_time(day_type, email):
    idx = ord(day_type) - ord('a') + 26
    if email[idx] == "1st Lunch - Cafeteria (10:15~10:45)":
        lunch_schedule = ["First", "Cafeteria"]
    elif email[idx] == "2nd Lunch - Cafeteria (10:50~11:20)":
        lunch_schedule = ["Second", "Cafeteria"]
    elif email[idx] == "1st Lunch - Conference Hall (10:15~10:45)":
        lunch_schedule = ["First", "Conference Hall"]
    elif email[idx] == "2nd Lunch - Conference Hall (10:50~11:20)":
        lunch_schedule = ["Second", "Conference Hall"]
    else:
        lunch_schedule = ["",""]
    #print(lunch_schedule)
    return lunch_schedule

def final_day(week_day):
    if (week_day == "Sunday"):
        return "Monday"
    elif (week_day == "Monday"):
        return "Tuesday"
    elif (week_day == "Tuesday"):
        return "Wednesday"
    elif (week_day == "Wednesday"):
        return "Thursday"
    elif (week_day == "Thursday"):
        return "Friday"
    
def execute(week_day, letter_day):
    store = file.Storage('/home/code/Desktop/scheduleBot/token.json')
    creds = store.get()
    if not creds or creds.invalid:
        flow = client.flow_from_clientsecrets('/home/code/Desktop/scheduleBot/credentials.json', SCOPES)
        creds = tools.run_flow(flow, store)
    service = build('sheets', 'v4', http=creds.authorize(Http()))

    sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId=SPREADSHEET_ID, range=formResponses).execute()
    values = result.get('values', [])

    for email in values: #scanning through all the emails
        if (email[30] == "Yes"): 
            #print(email)
            rec_email = email[1] #kevin's og code says index 0, but isn't it 1?
            name_blocks = block_names(letter_day, email)
            link_blocks = block_links(letter_day, email)
            lunch = lunch_location_time(letter_day, email)
            clubs = club_or_adv(week_day, email)
            #print(link_blocks)
            #print(len(link_blocks))
            block1_name = name_blocks[0]
            block2_name = name_blocks[1]
            block3_name = name_blocks[2]
            block4_name = name_blocks[3]
            #print(block1_name + " " + block2_name + " " + block3_name + " " + block4_name)
            block1_link = link_blocks[0]
            block2_link = link_blocks[1]
            block3_link = link_blocks[2]
            block4_link = link_blocks[3]
            #print(block1_link + " " + block2_link + " " + block3_link + " " + block4_link)
            lunch_block = lunch[0]
            lunch_location = lunch[1]

            club_title = clubs[0]
            club_link = clubs[1]
            #print(lunch_block + lunch_location)
            #for normal days (add links and stuff later)
            return_day = final_day(week_day)
            sbj = email[2] + "'s schedule for " + return_day + " " + letter_day + " day"
            #CHANGE SCHEDULE BASED ON PDF
            if lunch_block == "First": 
                class_msg = "8:00 ~ 9:21: "  + block1_name + " " + block1_link + "\n\n"
                class_msg = class_msg + "9:28 ~ 10:10: " + block2_name + " " + block2_link + "\n\n"
                class_msg = class_msg + "10:15 ~ 10:45: Lunch @ " + lunch_location + "\n\n"  
                class_msg = class_msg + "10:50~11:20: " + block2_name + " " + block2_link + "\n\n"
                class_msg = class_msg + "11:30 ~ 12:08: " + club_title + " " + club_link + "\n\n"
                class_msg = class_msg + "12:13 ~ 1:33: " + block3_name + " " + block3_link + "\n\n"
                class_msg = class_msg + "1:40 ~ 3:00: " + block4_name + " " + block4_link + "\n"
            elif lunch_block == "Second": 
                class_msg = "8:00 ~ 9:21: "  + block1_name + " " + block1_link + "\n\n"
                class_msg = class_msg + "9:28 ~ 10:48: " + block2_name + " " + block2_link + "\n\n"
                class_msg = class_msg + "10:50 ~ 11:25: Lunch @ " + lunch_location + "\n\n"  
                class_msg = class_msg + "11:30 ~ 12:08: " + club_title + " " + club_link + "\n\n"
                class_msg = class_msg + "12:13 ~ 1:33: " + block3_name + " " + block3_link + "\n\n"
                class_msg = class_msg + "1:40 ~ 3:00: " + block4_name + " " + block4_link + "\n"
            msg = class_msg
            #print(class_msg + "\n")
            
            if letter_day != 'x':
                
                server = smtplib.SMTP("smtp.googlemail.com:587")
                server.ehlo()   
                server.starttls()
                server.login(mail_user, password)
                text = MIMEText("Hello "+email[2]+",\n"+"Here's your schedule for tomorrow \n\n"+msg+"\n\nCurrently "+str(len(values))+" people are using Schedulebot."+"\n\nThank you for using Schedulebot\nCreated by Young Kim and Paul Ericksen.\n\nIf you wish to unsubscribe, fill out No for the last check box.\nhttps://www.kis.support/schedulebot")
                #print (text)
                text['Subject'] = sbj
                text['To'] = rec_email
                text['From'] = mail_user
                server.sendmail(mail_user,rec_email,text.as_string())
                #rec_email
                server.quit()
                sleep(3)
                #break
                print("sent")
            
def main():
    current_day = strftime("%A", localtime())
    current_date = strftime("%j", localtime())
    current_date = current_date.lstrip("0")
    current_time = int(strftime("%H", localtime()))
    #print(current_day, " ", current_time, " ", current_date)

    current_letter_day = finding_letter_day(current_date)
    execute(current_day, current_letter_day)

main()


