import speech_recognition #Importing speech recognition package from google api
import webbrowser
import pyaudio
import wolframalpha #To calculate strings into formula, its a website which provides api, 100 times per day
import requests
import re
import json
from playsound import playsound
from win32com.client import Dispatch
import sys
import os  #To save/open files
from time import *
from weather import Weather, Unit #API for weather

#Text-to-Speech Function
def TtSfunc(audio):
    for line in audio.splitlines():
        Dispatch('SAPI.SpVoice').Speak(audio)

#To check internet connection
def connected_to_internet(url='https://www.google.com/'):
    try:
        _ = requests.get(url, timeout = 1)
        return True
         
    except requests.ConnectionError:
        return False

if connected_to_internet():
    pass

else:
    TtSfunc('Check internet connection!')
    TtSfunc('Turning OFF, Goodbye!')
    playsound('close.wav')
    sys.exit()

#Speech-to-Text Initial Processes
with speech_recognition.Microphone() as source:
    speech_recognition.Recognizer().pause_threshold = 0.8
    speech_recognition.Recognizer().operation_timeout = 1
    speech_recognition.Recognizer().adjust_for_ambient_noise(source, duration = 1)
    speech_recognition.Recognizer().dynamic_energy_threshold = True
    TtSfunc('What can I do for you?')

#Speech-to-Text Function
def StTfunc():
    playsound('listen.wav')

    with speech_recognition.Microphone() as source:
        audio = speech_recognition.Recognizer().listen(source, timeout = 5)

    try:
        command = speech_recognition.Recognizer().recognize_google(audio, language = 'en-US').lower()        

    #Message when internet is not available
    except speech_recognition.RequestError:
        TtSfunc('Check internet connection!')
        sys.exit()

    #Message when no input is detected
    except speech_recognition.WaitTimeoutError:
        TtSfunc('I\'m listening to you, please speak!')
        command = StTfunc()

    #Loop back when command is not understandable
    except speech_recognition.UnknownValueError:
        TtSfunc('I didn\'t hear that, please speak again!')
        command = StTfunc()

    return command

#Execution of commands
def assistant(command):
    if 'open' in command:
        reg_ex = re.search('open (.*)', command)
        if reg_ex:
            domain = reg_ex.group(1)
            url = 'https://www.'+ domain + '.com/'
            webbrowser.open(url)
            TtSfunc('Opening ' + domain)
        else:
            pass

    elif 'search for' in command:
        reg_ex = re.search('search for (.*)', command)
        if reg_ex:
            query = reg_ex.group(1)
            url = 'https://www.google.co.in/search?q=' + query
            webbrowser.open(url)
            TtSfunc('Searching for ' + query)
        else:
            pass

    elif 'google for' in command:
        reg_ex = re.search('google for (.*)', command)
        if reg_ex:
            query = reg_ex.group(1)
            url = 'https://www.google.co.in/search?q=' + query
            webbrowser.open(url)
            TtSfunc('Googling for ' + query)
        else:
            pass

    elif 'search google for' in command:
        reg_ex = re.search('search google for (.*)', command)
        if reg_ex:
            query = reg_ex.group(1)
            url = 'https://www.google.co.in/search?q=' + query
            webbrowser.open(url)
            TtSfunc('Searching google for ' + query)
        else:
            pass

    elif 'search videos for' in command:
        reg_ex = re.search('search videos for (.*)', command)
        if reg_ex:
            query = reg_ex.group(1)
            url = 'https://www.youtube.com/results?search_query=' + query
            webbrowser.open(url)
            TtSfunc('Searching videos for ' + query)
        else:
            pass

    elif 'play' in command:
        reg_ex = re.search('play (.*)', command)
        if reg_ex:
            query = reg_ex.group(1)
            url = 'https://www.youtube.com/results?search_query=' + query
            webbrowser.open(url)
            TtSfunc('Playing ' + query)
        else:
            pass

    elif 'search wikipedia for' in command:
        reg_ex = re.search('search wikipedia for (.*)', command)
        if reg_ex:
            query = reg_ex.group(1)
            url = 'https://en.wikipedia.org/wiki/' + query
            webbrowser.open(url)
            TtSfunc('Searching wikipedia for ' + query)
        else:
            pass

    elif 'search meaning of' in command:
        reg_ex = re.search('search meaning of (.*)', command)
        if reg_ex:
            query = reg_ex.group(1)
            url = 'https://www.google.co.in/search?q=' + query + '%20meaning'
            webbrowser.open(url)
            TtSfunc('Telling you the meaning of ' + query)
        else:
            pass

    elif 'start' in command:
        reg_ex = re.search('start (.*)', command)
        if reg_ex:
            exe = reg_ex.group(1)
            if 'chrome' in exe:
                os.system("start chrome.exe")
                TtSfunc('Starting ' + exe)
            elif 'firefox' in exe or 'mozilla' in exe or 'mozilla firefox' in exe:
                os.system("start firefox.exe")
                TtSfunc('Starting ' + exe)
            elif 'sublime' in exe:
                os.system("start sublime_text.exe")
                TtSfunc('Starting ' + exe)
            elif 'file manager' in exe:
                os.system("start explorer.exe")
                TtSfunc('Starting ' + exe)
            elif 'word' in exe:
                os.system("start winword.exe")
                TtSfunc('Starting ' + exe)
            elif 'excel' in exe:
                os.system("start excel.exe")
                TtSfunc('Starting ' + exe)
            elif 'outlook' in exe:
                os.system("start outlook.exe")
                TtSfunc('Starting ' + exe)
            elif 'publisher' in exe:
                os.system("start mspub.exe")
                TtSfunc('Starting ' + exe)
            elif 'powerpoint' in exe:
                os.system("start powerpnt.exe")
                TtSfunc('Starting ' + exe)
            elif 'onenote' in exe:
                os.system("start onenote.exe")
                TtSfunc('Starting ' + exe)
            elif 'access' in exe:
                os.system("start msaccess.exe")
                TtSfunc('Starting ' + exe)
            else:
                TtSfunc('Application not supported')
                pass

    elif 'weather in' in command:
        reg_ex = re.search('weather in (.*)', command)
        if reg_ex:
            city = reg_ex.group(1)
            weather = Weather(unit = Unit.CELSIUS)
           location = weather.lookup_by_location(city)
           condition = location.condition
           TtSfunc(condition)
        else:
            TtSfunc('Oops! I couldn\'t find weather of ' + city)

    elif 'tell me a joke' in command:
        res = requests.get('https://icanhazdadjoke.com/', headers={"Accept":"application/json"})
        if res.status_code == requests.codes.ok:
            TtSfunc(str(res.json()['joke']))
            playsound('joke.wav')
        else:
            TtSfunc('Oops! I ran out of jokes')

    elif 'where is' in command:
        reg_ex = re.search('where is (.*)', command)
        if reg_ex:
            place = reg_ex.group(1)
            url = 'https://www.google.com/maps/place/' + place + '/&amp;'
            webbrowser.open(url)
            TtSfunc('Hold on, I\'ll show you where ' + place + ' is')
        else:
            pass

    elif 'calculate' in command:
        reg_ex = re.search('calculate (.*)', command)
        if reg_ex:
            app_id = '7QV8LG-VKEJA3YPEE'
            client = wolframalpha.Client(app_id)
            expr = reg_ex.group(1)
            res = client.query(expr)
            answer = next(res.results).text
            TtSfunc("The answer is " + answer)
        else:
            pass

    elif 'shutdown' in command:
        TtSfunc('Okay, shutting down!')
        os.system("shutdown /s")
        sys.exit()

    elif 'restart' in command:
        TtSfunc('Okay, restarting!')
        os.system("shutdown /r")
        sys.exit()

    elif 'log off' in command:
        TtSfunc('Okay, logging off!')
        os.system("shutdown /l")
        sys.exit()

    elif 'hibernate' in command:
        TtSfunc('Okay, hibernating!')
        os.system("shutdown /h")
        sys.exit()

    elif 'what\'s the time' in command or 'time please' in command:
        clock = time.strftime('%I:%M %p', localtime())
        TtSfunc('Time is : ' + clock)

    elif 'what\'s the date' in command or 'date please' in command:
        date = time.strftime('%A, %B\' %Y', localtime())
        TtSfunc('Date is : ' + date)

    elif 'close' in command or 'bye' in command or 'sleep' in command:
        TtSfunc('Okay! Goodbye...')
        playsound('close.wav')
        sys.exit()

    elif 'what are you doing' in command:
        TtSfunc('Responding to your commands')

    elif 'what\'s up' in command:
        TtSfunc('Fine!')

    elif 'how are you' in command:
        TtSfunc('I\'m good!')

    elif 'what can you do' in command:
        TtSfunc('I can help you with many things. You can ask me to search anything on web, open any website, start any application and many more!')

    elif 'who developed you' in command or 'who created you' in command:
        TtSfunc('I\'ve been developed by Mr. Harshit Gupta')

    elif 'tell me about yourself' in command or 'who are you' in command:
        TtSfunc('My name is Alex. I\'m your virtual assistant.')

    elif 'hey dude' in command:
        TtSfunc('Hey there!')

    elif 'crazy' in command or 'mental' in command:
        TtSfunc('Well, there are 2 mental asylums in India')

    elif 'you are so cool' in command or 'you are so good' in command:
        TtSfunc('Not more than you!')

    elif 'ok' in command or 'good' in command or 'excellent' in command or 'thank you' in command:
        TtSfunc('Hope you like it!')

    else:
        reg_ex = re.search('(.*)', command)
        if reg_ex:
            query = reg_ex.group(1)
            url = query
            TtSfunc('I can search the web for you, Do you want to continue?')
            ans = StTfunc()
            if 'yes' in ans  or 'yeah' in ans:
                webbrowser.open(url)
                TtSfunc('Searching for ' + url)
            elif 'no' in ans or 'nope' in ans:
                TtSfunc('Okay, fine!')
                pass

#Loop to continue executing commands
while True:
    assistant(StTfunc())
