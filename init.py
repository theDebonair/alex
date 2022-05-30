import speech_recognition
from win32com.client import Dispatch
import os
import requests
from playsound import playsound

def TtS(audio):
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
    TtS('Check internet connection!')
    TtS('Turning OFF, Goodbye!')
    playsound('close.wav')
    sys.exit()

#Speech-to-Text Initial Processes
with speech_recognition.Microphone() as source:
    speech_recognition.Recognizer().pause_threshold = 0.8
    speech_recognition.Recognizer().adjust_for_ambient_noise(source, duration = 1)
    speech_recognition.Recognizer().dynamic_energy_threshold = True
    TtS('Running')
    playsound('listen.wav')

def StT():
     with speech_recognition.Microphone() as source:
         audio = speech_recognition.Recognizer().listen(source)

         try:
             command = speech_recognition.Recognizer().recognize_google(audio, language = 'en-US').lower()

         except:
             command = StT()
            
         return command

#This function will open 'engine.py' script when command is spoken
def assistant(command):
    if 'hey alex' in command or 'listen alex' in command or 'alex' in command:
        os.system('gui.py')
        TtS('Initializing')

    else:
        pass

while True:
    assistant(StT())
