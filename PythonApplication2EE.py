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
password = "distcodegwlm"
print("started")
isSent = True
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

        #print("Sending soon")
        if email[11] == "Yes":
            #print(rec_email)
            
            if email[12] == "High School":
                first = "Highschool Lunch"
                second = "Advisory or Club"
            elif email[12] == "Middle School":
                first = "Advisory or Club"
                second = "Middleschool Lunch"
            if dayName == "a":
                sbj = email[3]+", "+email[4]+", "+email[5]+", "+email[6]
                msg = "8:00 - 9:21: " + email[3]+"\n" + "9:28 - 10:48: " + email[4]+"\n" + "10:53 - 11:25: " + first+ "\n" + "11:30 - 12:08: " + second+"\n" + "12:13 - 1:33: " + email[5]+"\n" + "1:40 - 3:00: " + email[6]
            elif dayName == "b":
                sbj = email[7]+", "+email[8]+", "+email[9]+", "+email[10]
                msg = "8:00 - 9:21: " + email[7]+"\n" + "9:28 - 10:48: " + email[8]+"\n" + "10:53 - 11:25: " + first+ "\n" + "11:30 - 12:08: " + second+"\n" + "12:13 - 1:33: " + email[9]+"\n" + "1:40 - 3:00: " + email[10]
            elif dayName == "c":
                sbj = email[6]+", "+email[5]+", "+email[4]+", "+email[3]
                msg = "8:00 - 9:21: " + email[6]+"\n" + "9:28 - 10:48: " + email[5]+"\n" + "10:53 - 11:25: " + first+ "\n" + "11:30 - 12:08: " + second+"\n" + "12:13 - 1:33: " + email[4]+"\n" + "1:40 - 3:00: " + email[3]
            elif dayName == "d":
                sbj = email[10]+", "+email[9]+", "+email[8]+", "+email[7]
                msg = "8:00 - 9:21: " + email[10]+"\n" + "9:28 - 10:48: " + email[9]+"\n" + "10:53 - 11:25: " + first+ "\n" + "11:30 - 12:08: " + second+"\n" + "12:13 - 1:33: " + email[8]+"\n" + "1:40 - 3:00: " + email[7]
            elif dayName == "e":
                sbj = email[4]+", "+email[3]+", "+email[6]+", "+email[5]
                msg = "8:00 - 9:21: " + email[4]+"\n" + "9:28 - 10:48: " + email[3]+"\n" + "10:53 - 11:25: " + first+ "\n" + "11:30 - 12:08: " + second+"\n" + "12:13 - 1:33: " + email[6]+"\n" + "1:40 - 3:00: " + email[5]
            elif dayName == "f":
                sbj = email[8]+", "+email[7]+", "+email[10]+", "+email[9]
                msg = "8:00 - 9:21: " + email[8]+"\n" + "9:28 - 10:48: " + email[7]+"\n" + "10:53 - 11:25: " + first+ "\n" + "11:30 - 12:08: " + second+"\n" + "12:13 - 1:33: " + email[10]+"\n" + "1:40 - 3:00: " + email[9]
            elif dayName == "Final":
                sbj = "You have Finals tomorrow."
                msg = "You have final exams. Check your Finals Schedule on the Fish Bowl windows."
            elif dayName == "PepRally4321":
                if email[12] == "High School":
                    sbj = email[6]+", "+email[5]+", "+email[4]+", "+email[3]
                    msg = "8:00 - 9:21: " + email[6]+"\n" + "9:28 - 10:48: " + email[5]+"\n" + "10:53 - 11:25: " + first+ "\n" + "11:30 - 12:08: " + second+"\n" + "12:13 - 1:10: " + email[4]+"\n" + "1:17 - 2:15: " + email[3] + "\n2:20 - 3:00: Winter Pep Rally"
                else:
                    sbj = email[6]+", "+email[5]+", "+email[4]+", "+email[3]
                    msg = "8:00 - 9:21: " + email[6]+"\n" + "9:28 - 10:48: " + email[5]+"\n" + "10:53 - 11:25: " + first+ "\n" + "11:30 - 12:08: " + second+"\n" + "12:13 - 1:33: " + email[4]+"\n" + "1:40 - 3:00: " + email[3]
            elif dayName == "Special12":
                sbj = email[3] +", "+email[4]
                msg = "8:00 - 9:21: " + email[3]+"\n" + "9:28 - 10:48: " + email[4]+"\n" + "10:53 - 11:25: " + first+ "\n" + "11:30 - 12:08: " + second+"\n" + "Parent Teacher Conferences"
            elif dayName == "Special34":
                sbj = email[5] +", "+email[6]
                msg = "8:00 - 9:21: " + email[5]+"\n" + "9:28 - 10:48: " + email[6]+"\n" + "10:53 - 11:25: " + first+ "\n" + "11:30 - 12:08: " + second+"\n" + "Parent Teacher Conferences"
            elif dayName == "EE1":
                if "23" in email[0]:
                    sbj = "Here is your EE Trip Itinerary"
                    msg = """

7:30 am Arrive at KIS
Write down one question about the trip that you can come back to at the end of the trip on the hike
8:15 am   
Depart KIS 
11 am
Groups arrive at Dongsan-ri beach
Eat packed lunches on the beach- teachers spread out among groups of students for broad supervision
12:45 - 4:00 pm
1st Rotation of Activities (Bike Ride, Surfing/Beach Time, Bio/Observation walk)
4:30 pm
Buses are loaded & depart for hotel 
5:00 pm: 
Arrive back at Naksan Condominium 
5:30 pm: 
Dinner
6:45 pm: 
Meet in the Building 5 Conference Hall for activity
7:45 pm:
Beach- play big team games
8:30 pm:
Back to rooms 
10:00 pm: 
Students are in rooms for 10 PM curfew.


                    """         
                
                elif "22" in email[0]:
                    sbj = "Here is your EE Trip Itinerary"
                    msg = """

8:15 a.m.
Load buses, take attendance, depart 

~9:15 a.m.
Rest stop - Anseong (bathroom & snacks)  15 minutes

10:30 a.m.
Arrive at Special Needs Home (Buses A and B) Group 1
Arrive at Pottery Museum (Buses C and D) Group 2

12:30 p.m.
Lunch (Student must pack a lunch)

1:30 p.m.
Arrive at Retirement Home (Buses C and D) Group 2
Arrive at Pottery Museum (Buses A and B) Group 1

3:30
Leave for Boramwon

4:00
Arrive at Boramwon - Check into rooms

5:00- 6:00 p.m.
Opening ceremonies/ Amphitheater

6:15 p.m.
Dinner

7:00 p.m. Nanta
8:00 p.m. Group 1 - Maze  to  Star Gazing/Color Run
9:00 p.m. Group 2 - Star Gazing/Color Run to  Maze

10:15 p.m.
Wash up & prepare for bed

11:00 p.m.
Lights Out!
        
                    """         
                elif "21" in email[0]:
                    sbj = "Here is your EE Trip Itinerary"
                    msg = """
6:15 am  
Arrive at KIS 

7:00 am  
Depart KIS (MUST depart on time to keep to schedule) 
There will be a rest stop, but NO PURCHASING please.

12:00 pm
Arrive @ hotel to drop off bags*
*Check to see if rooms are ready
12:45 pm
Meet your advisor and walk up to Aikwangwon!

4:30 pm: 
Meet your advisor and walk back to the hotel

5:00 pm: 
Dinner or Activities (depending on your group)

7:00 pm: 
Dinner or Activities (depending on your group)

9:00 pm: 
Meet with your advisor to talk about the day

9:30 pm: 
In rooms

10:00
Lights out, quiet sleepy time.           
                    """            
                elif "20" in email[0]:
                    sbj = "Here is your EE Trip Itinerary"
                    msg = """
8:15 am  
Depart KIS 

11:15
Buses Arrive At Manhae Hostel
Students 
Check in rooms
Eat lunch

Afternoon Rotations
Group A: Whitewater Rafting
Group B: Service Learning
Group C: Ulsanbawi Summit Hike  
Group D: Iron Way Climb 

5:00 pm: 
Dinner @ Bibimbap Restaurant in Seoraksan Park

6:30 pm: 
Visit Naksan Beach. Return to Manhae hostel in the event of extreme weather conditions.

9:00 pm: 
Return to Donguk University Youth Hostel - night snack

10:00 pm: 
In rooms/Curfew 
           
                    """
                elif "24" or "25" or "26" or "27" in email[0]:
                    sbj = email[4]+", "+email[3]+", "+email[6]+", "+email[5]
                    msg = "8:00 - 9:21: " + email[4]+"\n" + "9:28 - 10:48: " + email[3]+"\n" + "10:53 - 11:25: " + first+ "\n" + "11:30 - 12:08: " + second+"\n" + "12:13 - 1:33: " + email[6]+"\n" + "1:40 - 3:00: " + email[5]
                else:
                    sbj = "Sorry Teachers, For the EE Trip days, schedule bot will only work for students"
                    msg = "I apologize for the inconvenience."
            elif dayName == "EE2":
                if "23" in email[0]:
                    sbj = "Here is your EE Trip Itinerary"
                    msg = """

8:00 am 
Breakfast
9:00 am - 12:00 pm
2nd rotation of activities (Bike Ride, Surfing/Beach Time, Bio/Observation walk)
12:00 pm - 1:00 pm 
Lunch
1:00 pm - 4:00 pm
3rd rotation of activities (Bike Ride, Surfing/Beach Time, Bio/Observation walk)
4:30 pm
Buses are loaded & depart for hotel 
5:45 pm: 
Beach BBQ
5:45 pm-6:30 pm 
Sand sculpture contest & results on the beach. 
Release for bbq dinner based on sandcastle results.
6: 30 pm
In shifts, students are eating BBQ at tables. 
3 staff members get food early & head to the beach to supervise. 
7:00-7:30 pm
On Naksan beach: games, activities, social time
7:45-8:30 pm
Debrief / values follow up in Building 5 Conference Hall
8:30 pm: 
Clean up rooms and get ready for bed. Chill time.
10:00 pm:
Curfew for students.   

                    """         
                
                elif "22" in email[0]:
                    sbj = "Here is your EE Trip Itinerary"
                    msg = """


6:30 a.m. 
WAKE UP CALL!

7:30 a.m.
EAT BREAKFAST

8:30 a.m-9:45
ROTATION #1
Group A - puddle paddle ( Boat Battles)
Group B - archery
Group C - adventure course

10:00 - 11:15
ROTATION #2
Group A - archery
Group B - adventure course
Group C - puddle paddle  ( Boat Battles)

11:30 a.m.
Lunch

12:30 p.m. - 1:45
ROTATION #3
Group A - adventure course
Group B - puddle paddle  ( Boat Battles)
Group C - archery

2:00 - 4:00
ROTATION #A
Group 1 - hiking 
Group 2 -Swimming /debrief/   Field Games 

4:00 - 6:00
ROTATION #B
Group 1 - Swimming /debrief/   Field Games 
Group 2 - hiking

6:15- 7:00 p.m.
Dinner

7:00 p.m.
Soccer Field free time 

8:00 p.m.
Night Hike

9:00 p.m.- 10:30pm
Bonfire (Talent Show )

10:30 p.m.
Wash up & prepare for bed

11:00 p.m.
Lights Out!

        
                    """         
                elif "21" in email[0]:
                    sbj = "Here is your EE Trip Itinerary"
                    msg = """

7:15 am 
Shower/dress for the day.
Please wear footwear for hiking
Pack water

8:00 am  
Time to walk up to Aikwangwon

8:20 am  
Breakfast at Aikwangwon (in advisory groups)

8:50 am  
Head to buses (for hike) or Manna Hall (depending on group)
9:00 am
Group 1 : Hiking to the lighthouse
Group 2 : Activity with residents

12:00 am  
Head to the restaurant
 
12:15 pm 
lunch @ restaurant
Wash your hands before you eat.
Sit anywhere within the designated area, though you cannot leave the restaurant.
Relax, swap stories, enjoy each other's company
Go to the bathroom before you leave

1:15 pm  
Head to buses (for hike) or Manna Hall (depending on group)

1:30 pm
Group 1 : Activity with residents
Group 2 : Hiking to the lighthouse

4:00pm
Groups head to the front of Aikwangwon for the beach

5:00 - 8:00 pm
Beach time
Dinner: Made by your loving teachers
Advisories will be called in turn for dinner, listen out so you don't miss out! And don't forget to wash your hands, please. :)


8:30 pm: 
Advisory debrief time

9:00 pm: 
All students to their rooms

10:00 pm: 
Lights out, get some much-needed rest, for an early start tomorrow
        
                    """            
                elif "20" in email[0]:
                    sbj = "Here is your EE Trip Itinerary"
                    msg = """

Morning
Breakfast @ Donguk University Youth Hostel
Students will make their lunch for the day

Morning Rotations
Group A: Iron Way Climb
Group B: Whitewater Rafting
Group C: Service Volunteer
Group D: Ulsanbawi Summit Hike 

Lunch 

Afternoon Rotations
Group A: Service Volunteer
Group B: Ulsanbawi Summit Hike
Group C: Iron Way Climb
Group D: Whitewater Rafting

Evening

5:00 pm:
Free Time

6:00 pm: 
Dinner at Hostel (Eat with Advisory, debrief and skit planning)

7:00 pm: 
Advisory Led Activities

8:30 pm: 
Campfire and Smore Cooking

10:00 pm: 
In rooms/Curfew 

           
                    """
                elif "24" or "25" or "26" or "27" in email[0]:
                    sbj = email[8]+", "+email[7]+", "+email[10]+", "+email[9]
                    msg = "8:00 - 9:21: " + email[8]+"\n" + "9:28 - 10:48: " + email[7]+"\n" + "10:53 - 11:25: " + first+ "\n" + "11:30 - 12:08: " + second+"\n" + "12:13 - 1:33: " + email[10]+"\n" + "1:40 - 3:00: " + email[9]
                else:
                    sbj = "Sorry Teachers, For the EE Trip days, schedule bot will only work for students"
                    msg = "I apologize for the inconvenience."
            elif dayName == "EE3":
                if "23" in email[0]:
                    sbj = "Here is your EE Trip Itinerary"
                    msg = """

8:00 am
Wake up
8:30 am
Breakfast
9:00 am
Bus departs Naksan Condo
10:00 am
Arrive @ Bangtae San Recreational Forest
10:15 am
Begin hike up the gravel drive to your groups first spot
11:30 - 11:45 am 
Depart for KIS
2:15 pm 
Arrive back @ KIS


                    """         
                
                elif "22" in email[0]:
                    sbj = "Here is your EE Trip Itinerary"
                    msg = """



7:00 a.m. 
WAKE UP CALL!
Pack, clean room, and be ready for breakfast by 7:00

8:00  a.m.
EAT BREAKFAST/Pick up lunch

8:30  a.m.
Check Out/clean rooms
Move luggage to dorm lobby

9:00 - 10:30a.m.
Closing Ceremony/ Students will write letters to themselves

10:30 a.m.
Lunch picnic/BBQ set-up/Amphitheater(Rain in cafe)

11:00 a.m.
Load buses and take attendance

11:30 a.m.
Depart back to KIS from ?
Use bathroom before leaving

12:30 p.m.
Quick Rest Stop - Guemwang (15 minutes)

2:00 p.m.
Arrive at KIS

                    """         
                elif "21" in email[0]:
                    sbj = "Here is your EE Trip Itinerary"
                    msg = """

6:45 am 
Wake up & wash

7:00am
Rooms will be checked to ensure all students are up.
Pack your things up and load them on the buses when your advisor tells you to.
Ensure your rooms are clean

7:45pm
Last walk up to Aikwangwon

7:50 - 8:15 am
Group 1: Wash hands and prepare lunch
Group 2:  Eat your breakfast

8:15 am - 8:30 am
Group 1: Eat your breakfast
Group 2: Wash hands and prepare your lunch

8:30 am - 9:15 
Last chance to spend some time with your new resident friends, and to say goodbyes.

9:30 am 
Depart for KIS
Rest stop (no purchasing again, please)

2:30 pm
Arrive at KIS 
Ensure you have ALL your belongings
Go home and rest, you will probably be quite tired
Think about how you handled the trip, how has it affected you? 
Enjoy your weekend. XD

        
                    """            
                elif "20" in email[0]:
                    sbj = "Here is your EE Trip Itinerary"
                    msg = """

Morning 
Breakfast @ Donguk University Youth Hostel
Students will make their lunch for the day

Morning
Group Schedule
Group A: Ulsanbawi Summit Hike
Group B: Iron Way Climb
Group C: Whitewater Rafting
Group D: Service Volunteer

11:30 am 
Depart for KIS
Lunch on The Bus

2:30 pm 
Arrive at KIS
Student Journal Reflection

           
                    """
                elif "24" or "25" or "26" or "27" in email[0]:
                    sbj = email[3]+", "+email[4]+", "+email[5]+", "+email[6]
                    msg = "8:00 - 9:21: " + email[3]+"\n" + "9:28 - 10:48: " + email[4]+"\n" + "10:53 - 11:25: " + first+ "\n" + "11:30 - 12:08: " + second+"\n" + "12:13 - 1:33: " + email[5]+"\n" + "1:40 - 3:00: " + email[6]
                else:
                    sbj = "Sorry Teachers, For the EE Trip days, schedule bot will only work for students"
                    msg = "I apologize for the inconvenience."
            if dayName != 'x':
                
                
                server = smtplib.SMTP("smtp.googlemail.com:587")
                server.ehlo()   
                server.starttls()
                server.login(mail_user, password)
                text = MIMEText("Hello "+email[2]+",\n"+"\n"+"This is your schedule for tomorrow\n\n"+msg+"\n\nCurrently "+str(len(values))+" people are using Schedule Bot"+"\n\nThank you for using Schedule bot\nCreated by Kevin Park and Paul Ericksen.\n\nIf you wish to unsubscribe, fill out No for the last check box.\nhttps://www.everyview.net/schedule-bot-form.html")
                text['Subject'] = sbj
                text['To'] = rec_email
                text['From'] = mail_user
                server.sendmail(mail_user,rec_email,text.as_string())
                
                server.quit()
                sleep(1)

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
"""
while True:
    currentTime = int(strftime("%H",localtime()))
    
    if currentTime == 20:
        if isSent == False:
            isSent = True
            date = strftime("%j",localtime())
            date = date.lstrip("0")
            whatDayIsIt()
            sleep(84600) #Sleep for 23 1/2 hours
            #print(currentTime)
    elif currentTime == 21:
        isSent = False
"""
        
    
if __name__ == '__main__':
    date = strftime("%j",localtime())

    date = date.lstrip("0")
    print(date)
    whatDayIsIt()

