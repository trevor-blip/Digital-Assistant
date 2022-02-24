from __future__ import print_function
import datetime
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import pyttsx3
import pytz
import speech_recognition as sr
import subprocess
import wikipedia
import webbrowser
import os
import smtplib
import random

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/calendar.readonly']
MONTHS = ['january', 'february', 'march', 'april', 'may', 'june', 'july', 'august', 'september', 'october', 'november',
          'december']
DAYS = ['monday', 'tuesday', 'wednesday', 'thursday', 'friday', 'saturday', 'sunday']
DAY_EXTENTIONS = ['rd', 'thur', 'st', 'nd']
engine = pyttsx3.init("sapi5")
voices = engine.getProperty("voices")
engine.setProperty('voice', voices[0].id)
MASTER = 'Trevor'



# speak function will pronounce the string which is passed to it
def speak(text):
    engine.say(text)
    engine.runAndWait()


def wishMe():
    hour = int(datetime.datetime.now().hour)

    if 0 <= hour <= 12:
        speak('Good Morning' + MASTER)
    elif 12 <= hour <= 18:
        speak('Good Afternoon' + MASTER)
    else:
        speak("Good Evening" + MASTER)


wishMe()


def takeCommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        r.adjust_for_ambient_noise(source)
        audio = r.listen(source)

    try:
        print('Recognizing....')
        query = r.recognize_google(audio, language='eng-in')
        print(f"user said: {query}\n")
    except Exception as e:
        speak("Say that again please")
        query = takeCommand()
    return query.lower()


def media():
    speak('ok sir')
    speak('starting required application')
    speak('what do you want me to play for you')
    k = takeCommand()
    speak('ok sir playing' + k + 'for you')
    os.startfile('C:\\Windows.old\\Users\\Rehobiam\\Music' + k + '.mp3')


def shutdown():
    speak('understood sir')
    speak('connecting to command prompt')
    speak('shutting down your computer')
    os.system('shutdown -s')


def gooffline():
    speak('ok sir')
    speak('closing all systems')
    speak('disconnecting to servers')
    speak('going offline')
    quit()


def sendEmail(to, content):
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.ehlo()
    server.starttls()
    email = 'gandanhamotrevor@gmail.com'
    password = 'hyperlink'
    server.login(email, password)
    speak('please enter the email address do you want to send to..')
    emailAddress = input("Enter the email address")
    speak('Sending an email to ' + emailAddress)
    server.sendmail(emailAddress, to, content)


def authenticate_google():
    """Shows basic usage of the Google Calendar API.
    Prints the start and name of the next 10 events on the user's calendar.
    """
    creds = None
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)

        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('calendar', 'v3', credentials=creds)

    return service


def get_events(day, service):
    # Call the Calendar API
    date = datetime.datetime.combine(day, datetime.datetime.min.time())
    end_date = datetime.datetime.combine(day, datetime.datetime.max.time())
    utc = pytz.UTC
    date = date.astimezone(utc)
    end_date = end_date.astimezone(utc)

    events_result = service.events().list(calendarId='primary', timeMin=date.isoformat(), timeMax=end_date.isoformat(),
                                          singleEvents=True,
                                          orderBy='startTime').execute()
    events = events_result.get('items', [])

    if not events:
        speak('No upcoming events found.')
    else:
        speak(f"You have {len(events)} events on this day.")

        for event in events:
            start = event['start'].get('dateTime', event['start'].get('date'))
            print(start, event['summary'])
            start_time = str(start.split("T")[1].split("-")[0])
            if int(start_time.split(":")[0]) < 12:
                start_time = start_time + "am"
            else:
                start_time = str(int(start_time.split(":")[0]) - 12) + start_time.split(":")[1]
                start_time = start_time + "pm"

            speak(event["summary"] + " at " + start_time)


def get_date(text):
    text = text.lower()
    today = datetime.date.today()

    if text.count("today") > 0:
        return today

    if text.count("tomorrow") > 0:
        return today + datetime.timedelta(1)

    day = -1
    day_of_week = -1
    month = -1
    year = today.year

    for word in text.split():
        if word in MONTHS:
            month = MONTHS.index(word) + 1
        elif word in DAYS:
            day_of_week = DAYS.index(word)
        elif word.isdigit():
            day = int(word)
        else:
            for ext in DAY_EXTENTIONS:
                found = word.find(ext)
                if found > 0:
                    try:
                        day = int(word[:found])
                    except:
                        pass

    # THE NEW PART STARTS HERE
    if month < today.month and month != -1:  # if the month mentioned is before the current month set the year to the next
        year = year + 1

    # This is slighlty different from the video but the correct version
    if month == -1 and day != -1:  # if we didn't find a month, but we have a day
        if day < today.day:
            month = today.month + 1
        else:
            month = today.month

    # if we only found a dta of the week
    if month == -1 and day == -1 and day_of_week != -1:
        current_day_of_week = today.weekday()
        dif = day_of_week - current_day_of_week

        if dif < 0:
            dif += 7
            if text.count("next") >= 1:
                dif += 7

        return today + datetime.timedelta(dif)

    if day != -1:  # FIXED FROM VIDEO
        return datetime.date(month=month, day=day, year=year)


def note(text):
    date = datetime.datetime.now()
    file_name = str(date).replace(":", "-") + "-note.txt"
    with open(file_name, "w") as f:
        f.write(text)

    subprocess.Popen(["notepad.exe", file_name])


def mainfunction():
    WAKE = 'wake up'
    SERVICE = authenticate_google()

    while True:
        print('Listening........')
        query = takeCommand()

        if query.count(WAKE) > 0:
            speak('i am ready')
            query = takeCommand()
            # Logic for executing tasks as per the query
            CALENDAR_STRS = ["what do i have", "do i have plans", "am i busy"]
            for phrase in CALENDAR_STRS:
                if phrase in query:
                    date = get_date(query)
                    if date:
                        get_events(date, SERVICE)
                    else:
                        speak("I didn't quite get that")

            NOTE_STRS = ["make a note", "write this down", "remember this", "type this"]
            for phrase in NOTE_STRS:
                if phrase in query:
                    speak("What would you like me to write down? ")
                    write_down = takeCommand()
                    note(write_down)
                    speak("I've made a note of that.")
                    mainfunction()

            wiki = ['wikipedia', 'wiki']
            for phrase in wiki:
                if phrase in query:
                    print('Searching Wikipedia')
                    query = query.replace("wikipedia", "")
                    results = wikipedia.summary(query, sentences=2)
                    print(results)
                    speak(results)
                    mainfunction()

            msc = ['play music', 'song']
            for phrase in msc:
                if phrase in query:
                    try:
                        media()
                    except:
                        speak('Cannot find match please try again')
                        media()

            name = ['what is your name', 'what do i call you', 'your name']
            for phrase in name:
                if phrase in query:
                    speak("my name is Gunz")
                    mainfunction()

            time = ['the time', 'time is it']
            for phrase in time:
                if phrase in query:
                    time = datetime.datetime.now().strftime('%H:%M:%S')
                    speak('the time is......' + time)
                    mainfunction()

            end = ['goodbye', 'kill program', 'close program', 'go offline', 'exit'
                                                                             '']
            for phrase in end:
                if phrase in query:
                    gooffline()

            grtn = ['hi', 'hey', 'whatsup', 'sup', 'you good']
            reply = ['how are you sir.....', 'i am great how is it sir']
            for phrase in grtn:
                if phrase in query.lower():
                    d = random.choice(reply)
                    speak(d)
                    mainfunction()

            shutdwn = ['switch off', 'shutdown']
            for phrase in shutdwn:
                if phrase in query:
                    shutdown()

            email = ['send email', 'email', 'send an email']
            for phrase in email:
                if phrase in query:
                    try:
                        speak("what should i send..")
                        content = takeCommand()
                        speak('please enter the email address you want to send to..')
                        to = input("Enter the email address: ")
                        sendEmail(to, content)
                        speak("email has been sent successfully")
                        mainfunction()
                    except Exception as e:
                        print(e)

            browse = ['search','google']
            for phrase in browse:
                if phrase in query:
                    speak('What do you want to search')
                    search = takeCommand()
                    url = 'https://google.com/search?q=' + search
                    chrome_path = "C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe"
                    webbrowser.register('chrome', None, webbrowser.BackgroundBrowser(chrome_path), 1)
                    webbrowser.get('chrome').open_new_tab(url)
                    print('Here is what i found on google' + search)
                    speak('Here is what i found on google' + search)

            loc = ['find location','location','map']
            for phrase in loc:
                if phrase in query:
                    speak('what is the location ')
                    location = takeCommand()
                    url = 'https://google.nl/maps/place/' + location + '/&amp;'
                    webbrowser.get().open(url)
                    print('here is the location of: ' + location)

        """
        
            elif 'open google' in query.lower():
                url = 'google.com'
                webbrowser.open(url)
    """

mainfunction()
