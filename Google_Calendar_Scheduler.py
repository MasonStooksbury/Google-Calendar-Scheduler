#! python3

#Google_Calendar_Scheduler.py - A program that adds events to Google Calendar
#                            - A program that logs into my work schedule (and my wife's) website, scrapes shift
#                                   data, scrubs it, and sends it to Google Calendar

########## VERSION HISTORY ###########
# Version 1:  Able to add an event to Google Calendar
# Version 2:  Able to add work or regular event (work events had hardcoded location),
#             Able to add a color to the event
#             Added COLORS dictionary
# Version 3:  Able to add work event with automated color, location, and title based on given name
# Version 4:  Able to add multiple events using a while loop
# Version 5:  Able to add multiple work events via an input file
# Version 6:  Able to add multiple work events for multiple users via input file
# Version 7:  Added abstracted "sendWorkEvent" function to easily send work events
#             Added "users" nested dictionary with all of the static user info (name, title, and color)
# Version 8:  Added abstracted "sendRegEvent" function to easily send regular events with no file
# Version 9:  Edited "event" dictionary to have different questions for work event
#             Added "infoCorrectPrompt" and "workInfoCorrectPrompt" functions to ask user if information
#                has been correctly entered
#             Added "joinTime" function to allow user to enter 4 consecutive digits without ":", handles
#                re-concatenation and returns a string
#             Added "predictETime" function to auto-fill the end-time for any given start-time whether
#                single or double shift (5 hours added for single, 4 for double)
#             Added "sendDouble" function to handle sending a more complex double shift
#             Refactored "Work Event, File" to be more clean and organized
#             Added "condatenate" function to allow user to enter a consecutive, 4-digit month-day combination
#                and will automatically determine the year and re-concatenate to "year-month-day" format
#             Added "datetime" element as "NOW" global to allow for automatic knowledge of the year
#             Edited "infoCorrectPrompt" and "workInfoCorrectPrompt" to display "condatenated" dates
#             Edited "event_details" dictionary to account for "condatenated" dates and "joinTime" times
#             Edited "sendRegEvent" and "sendWorkEvent" to account for "condatenated" dates and "joinTime" times
#             Edited "Work Event, No File" to account for predicted end-times
# Version 10: Added Function descriptions
# Version 11: Added ability to only type 'M', 'P', and 'D' for 'Mason', 'Paige', and 'Double' respectively
# Version 12: Edited "Work Event, File" to close the open file object
# Version 13: CMD Compliant
# Version 14: Edited "Work Event, File" to have FULL_PATH when looking for 'work.txt'
# Version 15: File is WITH Compliant
# Version 16: FULL OVERHAUL
#             Able to scrape date and times from website
#             Able to check for next week schedule
#             Added 'chill' function to shorten implicit wait Selenium command
#             Added 'login' function to handle logging in user
#             Added 'grabWebText' function to grab the html from the schedule screen
#             Added 'isNextWeek' function to return Boolean if "Next Week" is found on the page
#             Added 'textUs' function to text us the status of our schedules upon completion or failure
#             Added 'sendErrorText' to text me if something fails
#             Added 'sendStartText' to text me when computer and script are up and running
#             Added 'loadNextWeek' to handle clicking of the Next Week radio button and view schedule button
#             Added 'stringToDate' to convert first date on website to datetime.date object
#             Added 'getWeekDate' to grab week start date from website
#             Added 'setupShiftDict' to match all week date strings with list of work times
#             Added 'changeTimeFormat' to change all times to HHMM and 24 hr format
#             Added 'sendShift' to send a single shift
#             Added 'sendAllShifts' to handle looping thru WEEK_SHIFTS to send them all
#             Added 'executeLazyProtocol' to handle calling all the necessary functions for each user
########################################

from __future__ import print_function
from apiclient import discovery
from httplib2 import Http
from oauth2client import file, client, tools
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from twilio.rest import Client

import datetime #Keep this so it always knows when it is
import webbrowser
import time
import re
import lxml.html
import pprint

NOW = datetime.datetime.now()
FULL_PATH = 'C:\\Users\\MasonStooksbury\\Documents\\PythonScripts\\Scripts\\Extra_Stuff'

SCOPES = 'https://www.googleapis.com/auth/calendar'
store = file.Storage('gca_storage.json')

creds = store.get()
if not creds or creds.invalid:
    flow = client.flow_from_clientsecrets(FULL_PATH + '\\gca_client_secret.json', SCOPES)
    creds = tools.run_flow(flow, store)

GCAL = discovery.build('calendar', 'v3', http=creds.authorize(Http()))
TIMEZONE = 'America/Chicago'
WORK_LOC = 'Olive Garden Italian Restaurant'

COLORS = {'Default': 0,
          'Lavender': 1,
          'Sage': 2,
          'Grape': 3,
          'Flamingo': 4,
          'Banana': 5,
          'Tangerine': 6,
          'Peacock': 7,
          'Graphite': 8,
          'Blueberry': 9,
          'Basil': 10,
          'Tomato': 11
          }

web_path = '''https://krowdweb.darden.com/krowd/prd/siteminder/login.asp?TYPE=33554433
&REALMOID=06-e5941905-f720-4b70-9f12-cc19d477a633&GUID=&SMAUTHREASON=0&METHOD=GET&S
MAGENTNAME=-SM-L%2bGohbnlQ1dzLhWEg1OHiag5gcFHwyRoqASf57cA84BeOIDbswwJn3XYpIi3fUDo&TARGET
=-SM-https%3a%2f%2fkrowd%2edarden%2ecom%2fwebcenter%2fportal%2fKrowd%2fhome'''

USERS = { 'M': {'title': 'Work (Mason)',
                'color': 'Default',
                'username': 'WEBSITEUSERNAME',
                'password': '~WEBSITEPASSWORD~',
                'accountSID': 'AC599418d69d5f17c9069097cbd946d0dc',
                'authToken': '5d9addcd6a67b5040f1b7b4f34446441',
                'twilio': 'TWILIOCELL#',
                'cell': 'CELL#'
                },
          'P': {'title': 'Work (Paige)',
                'color': 'Tomato',
                'username': 'WEBSITEUSERNAME',
                'password': 'WEBSITEPASSWORD',
                'accountSID': 'ACcb98c2272212cf5d268fdfc88ddab73d',
                'authToken': 'd0af5fdd3d04ea3d88494ba6d07bd5c3',
                'twilio': 'TWILIOCELL#',
                'cell': 'CELL#'
                }
          }

WEEK = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday']



# Shortened version of the command with "time" argument
# Returns: Nothing
def chill(time):
    browser.implicitly_wait(time)
    return

# Opens a FireFox webdriver, navigates to user schedule.
# Returns: Nothing
def login(user):

    browser.get(web_path)
    
    fill_out_user = browser.find_element_by_id('user')
    fill_out_user.send_keys(USERS[user]['username'])

    fill_out_pass = browser.find_element_by_id('password')
    fill_out_pass.send_keys(USERS[user]['password'])

    fill_out_pass.submit()

    chill(30)

    # We don't actually need this. But it works best with 'chill' so we will keep it.
    browser.find_element_by_xpath("/html/body/div[2]/div/form/div/div[3]/div[1]/div/div[3]/div[2]/div/a").click()

        #I'm forcing this one because sometimes it's too jumpy even with 'chill'
        #time.sleep(8)
    
        #browser.find_element_by_xpath("/html/body/div[3]/div/div/div[2]/div/div/div[2]/div[14]/a[2]").click()

        #chill(7)

    browser.get("https://pconweb.darden.com/applications/schedules/schedule.asp?rst=true")
    
    chill(30)
    return

# Grabs all of its text
# Returns: String of all text on page
def grabWebText():

    all_text = browser.find_element_by_tag_name('body')
    
    return all_text.text

# Determines whether or not "Next Week" is on the page
# Returns: Boolean
def isNextWeek(string):
    next_week_regex = re.compile(r'\bNext Week\b')
    match = next_week_regex.search(string)

    if match.group() == 'Next Week':
        return True
    else:
        return False

# Uses Twilio to text us an update
# Returns: Nothing
def textUs(up_or_not, user):

    twilioCli = Client(USERS[user]['accountSID'], USERS[user]['authToken'])

    if up_or_not:
        message = "Schedules are posted in your Google Calendar! :D"

    else:
        message = "Schedules are not done... :("
    
    twilioCli.messages.create(
        to=USERS[user]['cell'],
        from_=USERS[user]['twilio'],
        body=message)
    return

# Sends me a text if something goes wrong
# Returns: Nothing
def sendErrorText(num):
    user = 'M'
    twilioCli = Client(USERS[user]['accountSID'], USERS[user]['authToken'])

    message = "ERROR ON ATTEMPT " + (num + 1) + ". I broke trying to refresh the page. I'm sorry :("
    
    twilioCli.messages.create(
        to=USERS[user]['cell'],
        from_=USERS[user]['twilio'],
        body=message)
    return

# Sends me a text that it's alive and looking for schedules
# Returns: Nothing
def sendStartText():
    user = 'M'
    twilioCli = Client(USERS[user]['accountSID'], USERS[user]['authToken'])

    message = "I'm awake, logged in, and about to look for schedules!"
    
    twilioCli.messages.create(
        to=USERS[user]['cell'],
        from_=USERS[user]['twilio'],
        body=message)
    return

# Clicks the 'Next Week' radio button and then the 'View Schedule' button
# Returns: Nothing
def loadNextWeek():
    chill(5)
    
    browser.find_element_by_xpath(next_week_radio_xpath).click()

    chill(2)

    browser.find_element_by_xpath(view_schedule_button).click()

    chill(5)

    return

# Accepts string date and converts to YYYY-MM-DD
# Returns: datetime object
def stringToDate(date):
    
    # Single day and month
    if date[1] == '/' and date[3] == '/':
        next_day = date[2:3]
        next_month = date[0:1]
        next_year = date[-4:]

    # Double month, single day
    elif date[2] == '/' and date[4] == '/':
        next_day = date[3:4]
        next_month = date[0:2]
        next_year = date[-4:]

    # Single month, double day
    elif date[1] == '/' and date[4] == '/':
        next_day = date[2:4]
        next_month = date[0:1]
        next_year = date[-4:]

    # Double month and day
    elif date[2] == '/' and date[5] == '/':
        next_day = date[3:5]
        next_month = date[0:2]
        next_year = date[-4:]

    new_date = datetime.date(int(next_year), int(next_month), int(next_day))
    return new_date

# Scrapes the start of the week date from website
# Returns: datetime object version of string date
def getWeekDate():
    
    search_area = root.xpath('/html/body/table[2]/tbody/tr[1]/th/text()[2]')
    regex = re.compile(r'(From) (\d(\d)?/\d(\d)?/\d{4})')
    match = regex.search(search_area[0])
    week_start_date = match.group(2)
    
    return stringToDate(week_start_date)

# Concatenates all date object strings with appropriate shift times and stores them in WEEK_SHIFTS
# Returns: Nothing
# External Modifications: WEEK_SHIFTS
def setupShiftDict(date):
    add_day = datetime.timedelta(1)
    
    WEEK_TIMES = {  'Monday' : root.xpath('.//table[2]//tr[3]//td[2]/span/text()'),
                    'Tuesday' : root.xpath('.//table[2]//tr[3]//td[3]/span/span/text()'),
                    'Wednesday' : root.xpath('.//table[2]//tr[3]//td[4]/span/span/text()'),
                    'Thursday' : root.xpath('.//table[2]//tr[3]//td[5]/span/span/text()'),
                    'Friday' : root.xpath('.//table[2]//tr[3]//td[6]/span/span/text()'),
                    'Saturday' : root.xpath('.//table[2]//tr[3]//td[7]/span/span/text()'),
                    'Sunday' : root.xpath('.//table[2]//tr[3]//td[8]/span/span/text()')
                 }

    for day in WEEK:
        WEEK_SHIFTS[str(date)] = WEEK_TIMES[day]
        date += add_day

    return

# Changes the format of the times in WEEK_SHIFTS to be HHMM and 24 hr format
# Returns: Nothing
# External Modifications: WEEK_SHIFTS dictionary values
def changeTimeFormat():
    
    for k, v in WEEK_SHIFTS.items():
        if len(v):
            # For evening times
            if v[0][2] != ':':
                new_hour = str(int(v[0][0:1]) + 12)
                WEEK_SHIFTS[k] = [new_hour + v[0][2:4]]
            else:
                WEEK_SHIFTS[k] = [v[0][0:2] + v[0][3:5]]

            # This try block accounts for Double-shift days
            try:
                if len(v) > 3:
                    if v[2] != 'Server':
                        new_hour = str(int(v[2][0:1]) + 12)
                        WEEK_SHIFTS[k].append(new_hour + v[2][2:4])
                        
                    # This handles Monday
                    elif v[2] == 'Server':
                        new_hour = str(int(v[3][0:1]) + 12)
                        WEEK_SHIFTS[k].append(new_hour + v[3][2:4])
            except:
                continue
            
    return

# Allows the user to enter only a start-time and will automatically calculate an estimate end-time.
#       Adds 5 hours for single shift, and 4 hours per shift for a double shift
# Returns: "time" String
def predictETime(t_type, time):
    if t_type == 'D':
        time = int(time) + 400
        return str(time)
    else:
        time = int(time) + 500
        return str(time)

# Reads in a string of 4 digits in hour-minute format (e.g. 1:30 PM = 1330) and adds
#       and adds a colon in between them
# Returns: "time" String
def joinTime(time):
    t_hour = time[:2]
    t_min = time[2:]
    time = t_hour + ':' + t_min
    return time

# Sends a single shift to Google Calendar
# Returns: Nothing
# External Modifications: Google Calendar
def sendShift(user, day, start, end):
    EVENT = {
                'summary': USERS[user]['title'],
                'start':  {'dateTime': day + 'T' + start + ':00', 'timeZone': TIMEZONE},
                'end':    {'dateTime': day + 'T' + end + ':00', 'timeZone': TIMEZONE},
                'location': WORK_LOC,
                'colorId': int(COLORS[USERS[user]['color']])
            }          
    e = GCAL.events().insert(calendarId='primary', sendNotifications=True, body=EVENT).execute()
    return

# Runs a 'for' loop on WEEK_SHIFTS to send all shifts
# Returns: Boolean if complete
# External Modifications: Google Calendar
def sendAllShifts(user):

    for k, v in WEEK_SHIFTS.items():
        if len(v):
            if len(v) > 1:
                sendShift(user, k, joinTime(v[0]), joinTime(predictETime('D', v[0])))
                sendShift(user, k, joinTime(v[1]), joinTime(predictETime('D', v[1])))
                
            else:
                sendShift(user, k, joinTime(v[0]), joinTime(predictETime('S', v[0])))

    return True

# Easy function to execute everything
# Returns: Nothing
# External Modifications: A LOT. Follow each function
def executeLazyProtocol(user):
    
    setupShiftDict(getWeekDate())

    #pprint.pprint(WEEK_SHIFTS)

    changeTimeFormat()

    #pprint.pprint(WEEK_SHIFTS)

    shift_status = sendAllShifts(user)

    textUs(shift_status, user)

    return





##### MAIN #####

sendStartText()

for person, stuff in USERS.items():
    WEEK_SHIFTS = {}
    
    browser = webdriver.Firefox()

    login(person)

    count = 0
    status = False


    # It will run until schedules are posted or it's been 3 hours with no luck.
    while not status or (count == 180):

        if count != 0:
            time.sleep(60)

        try:
            browser.refresh()
            chill(3)
            status = isNextWeek(grabWebText())
        except:
            sendErrorText(count)

        if status:
            break

        count += 1
    # END WHILE #########################
            

    if status:
        loadNextWeek()

        root = lxml.html.fromstring(browser.page_source)

        executeLazyProtocol(person)

        browser.quit()

    else:
        #textUs(isNextWeek(grabWebText()), 'Mason')
        user = 'M'
        twilioCli = Client(USERS[user]['accountSID'], USERS[user]['authToken'])
        message = "I tried for " + ((count + 1)/12) + " hours and they still haven't posted. I'm sorry :("

        twilioCli.messages.create(
                to=USERS[user]['cell'],
                from_=USERS[user]['twilio'],
                body=message)

##### END MAIN #####




