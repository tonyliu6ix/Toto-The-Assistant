import pyowm
import pyautogui
import random
import webbrowser
import datetime
import datefinder
import calendar
import pickle
import os
import time
import pyttsx3
import speech_recognition as sr
import pytz
import subprocess
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

SCOPE = ['https://www.googleapis.com/auth/calendar.readonly']
r, mic = sr.Recognizer(), sr.Microphone()


def speak(string):
    engine = pyttsx3.init()
    engine.setProperty('voice', engine.getProperty('voices')[0].id)
    engine.setProperty('rate', engine.getProperty('rate'))
    engine.setProperty('volume', engine.getProperty('volume') + 1.50)
    engine.say(string)
    engine.runAndWait()


def get_audio_data():
    with mic as source:
        r.adjust_for_ambient_noise(source)
        audio = r.listen(source)
        voice_data = ""
        try:
            voice_data = r.recognize_google(audio)
            print(voice_data)
        except sr.RequestError:
            print("API unavailable")
        except sr.UnknownValueError:
            print("Unable to recognize speech")
    return voice_data.lower()


def auth_calendar_api():
    creds = None
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPE)
            creds = flow.run_local_server(port=0)
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)
    service = build('calendar', 'v3', credentials=creds)
    return service


def get_twh_clock_time(hour, minute):
    if int(hour) < 12:
        if int(hour) == 0:
            twh_clock_time = '12:{} {}'.format(minute, 'am')
        else:
            twh_clock_time = '{}:{} {}'.format(hour, minute, 'am')
    else:
        if int(hour) == 12:
            twh_clock_time = '12:{} {}'.format(minute, 'pm')
        else:
            twh_clock_time = '{}:{} {}'.format(int(hour) - 12, minute, 'pm')
    return twh_clock_time


def get_curr_time():
    hour, minute = time.ctime().split()[3].split(':')[0], time.ctime().split()[3].split(':')[1]
    return get_twh_clock_time(hour, minute)


def get_events(day, service):
    date = datetime.datetime.combine(day, datetime.datetime.min.time()).astimezone(pytz.UTC)
    end_date = datetime.datetime.combine(day, datetime.datetime.max.time()).astimezone(pytz.UTC)
    events = service.events().list(
        calendarId='primary', timeMin=date.isoformat(), timeMax=end_date.isoformat(),
        singleEvents=True, orderBy='startTime'
    ).execute().get('items', [])
    if not events:
        speak("No upcoming events on this day")
    else:
        speak("You have {} events on this day".format(len(events))) if len(events) != 1 else \
            speak("You have 1 event on this day")
        for event in events:
            start = event['start'].get('dateTime', event['start'].get('date'))
            print(start, event['summary'])
            tfh_start_time = start.split("T")[1].split("-")[0].split(":")
            hour_start, minute_start = tfh_start_time[0], tfh_start_time[1]
            speak("You have {} at {}".format(event["summary"], get_twh_clock_time(hour_start, minute_start)))


def get_date(string):
    new_string = string.lower()
    today = datetime.date.today()
    lst_of_dates = []
    if "today" in new_string:
        lst_of_dates.append(today)
    if "tomorrow" in new_string:
        if today == datetime.date(today.year, today.month, calendar.monthrange(today.year, today.month)[1]):
            if today.month != 12:
                lst_of_dates.append(datetime.date(today.year, today.month + 1, 1))
            else:
                lst_of_dates.append(datetime.date(today.year + 1, 1, 1))
        else:
            lst_of_dates.append(datetime.date(today.year, today.month, today.day + 1))
    for dt_obj in list(datefinder.find_dates(new_string)):
        lst_of_dates.append(dt_obj.date())
    return lst_of_dates


def make_note(string):
    file_name = str(datetime.datetime.now()).replace(":", "-") + "_Created-By-Toto.txt"
    with open(file_name, "w") as f:
        f.write(string)
    subprocess.Popen(["notepad.exe", file_name])


def main():
    WAKEWORD = "toto"
    SERVICE = auth_calendar_api()
    speak("Hello, I'm Toto the virtual assistant, say 'Toto' for my help")
    utterance = get_audio_data()
    while WAKEWORD not in utterance:
        speak("Listening for my name, Toto")
        utterance = get_audio_data()
    speak("I am here, say exit for me to leave")
    asked_first_q = False
    while True:
        time.sleep(1.5)
        if asked_first_q:
            speak("What else would you like me to do?")
        else:
            speak("What is the first thing you would like me to do?")
        utterance = get_audio_data()
        if utterance != '':
            asked_first_q = True
        if utterance == "exit":
            speak("See you later alligator")
            break
        if 'time' in utterance:
            speak("It is " + get_curr_time())
        FORECAST_WORDS = ["temperature", "degrees", "hot", "cold", "weather", "humid"]
        if any(word in utterance for word in FORECAST_WORDS):
            speak('Which city do you live in?')
            loc_weather = ''
            loc = get_audio_data()
            while not loc_weather:
                try:
                    loc_weather = pyowm.OWM('22e94cadcf00b1c15795c0f46995db15').weather_at_place(loc).get_weather()
                except (
                pyowm.exceptions.api_response_error.NotFoundError, pyowm.exceptions.api_call_error.APICallError):
                    speak("Say your city again")
                    loc = get_audio_data()
            speak('Celsius or Fahrenheit?')
            metric = get_audio_data()
            while metric not in ['celsius', 'fahrenheit']:
                speak('Try again')
                metric = get_audio_data()
            temp, humidity = loc_weather.get_temperature(metric)['temp'], loc_weather.get_humidity()
            speak('It is {} degrees {} with a humidity of {} % in {}'.format(temp, metric, humidity, loc))
        SCREENSHOT_WORDS = ["capture", "my screen", "screenshot"]
        if any(word in utterance for word in SCREENSHOT_WORDS):
            ss = pyautogui.screenshot()
            speak("Screenshot saved")
            ss.save('C:\\Users\\Tony\\Desktop\\{}.png'.format(str(datetime.datetime.now()).replace(":", "-")))
        GOOGLE_SEARCH_WORDS = ['search', 'look up', 'google search', 'google something', 'google this']
        if any(word in utterance for word in GOOGLE_SEARCH_WORDS):
            speak("What do you want me to look up")
            search_term = get_audio_data()
            while not search_term:
                speak('I did not get that, try again')
                search_term = get_audio_data()
            webbrowser.get().open('https://google.com/search?q=' + search_term)
            speak("I found this for " + search_term)
        GOOGLE_MAPS_WORDS = ['location', 'place', 'map']
        if any(word in utterance for word in GOOGLE_MAPS_WORDS):
            speak("What location do you want?")
            loc = get_audio_data()
            while not loc:
                speak('Try again')
                loc = get_audio_data()
            webbrowser.get().open('https://google.nl/maps/place/' + loc + '/&amp;')
            speak("I found {} on Google Maps".format(loc))
        CALENDAR_WORDS = ["events", "plans", "have anything", "duties", "busy", "have", "schedule"]
        if any(word in utterance for word in CALENDAR_WORDS):
            dates = get_date(utterance)
            if dates:
                for date in dates:
                    get_events(date, SERVICE)
            else:
                speak("Please try again")
        NOTE_WORDS = ["a note", "write this ", "jot this", "jot down", "write down"]
        if any(word in utterance for word in NOTE_WORDS):
            speak("What do you want to write down?")
            make_note(get_audio_data())
            speak("I've made the note")


if __name__ == '__main__':
    main()
