import speech_recognition as sr
import os
from playsound import playsound
import webbrowser
import random
import wikipedia
import pyttsx3
import smtplib
import pyautogui
import psutil
import pyjokes
import wolframalpha
from selenium import webdriver
from gtts import gTTS
import google
import subprocess
import pickle
import json
import os.path
import datetime
import winshell
import feedparser
import time
import pytz
import shutil
import requests
from urllib.request import urlopen
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import time
from gtts import gTTS
from ecapture import ecapture as ec
from bs4 import BeautifulSoup
import win32com.client as wincl
from urllib.request import urlopen

def speak(text):
    tts = gTTS(text=text, lang="en-us")
    filename = "voice.mp3"
    tts.save(filename)
    playsound(filename)


engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[0].id)
engine.setProperty('rate', 100)
engine.setProperty('volume', 300)
engine.runAndWait()
speech = sr.Recognizer()

greeting_dict = {'hello': 'hello', 'hi': 'hi'}
open_launch_dict = {'open': 'open', 'launch': 'launch'}
google_searches_dict = {'what': 'what', 'which': 'which', 'when': 'when', 'where': 'where', 'search': 'search',
                        'who': 'who', 'whom': 'whom', 'whose': 'whose', 'why': 'why', 'whether': 'whether'}
social_media_dict = {'facebook': 'https://www.facebook.com/', 'instagram': 'https://www.instagram.com/',
                     'twitter': 'https://twitter.com/explore', ' whatsapp': 'https://www.whatsapp.com/'}
youtube_search_dict = {'play': 'play'}
youtube_song_dict = {'gane': 'gane', 'song': 'song'}

mp3_thankyou_list = ['MP3\welcome.mp3']
mp3_listening_problem_list = ['MP3\pardon.mp3', 'MP3\sorry.mp3']
mp3_struggling_list = ['MP3\samaj.mp3', 'MP3\struggle.mp3']
mp3_bye = ['MP3\goodbye.mp3', 'MP3\ciao.mp3', ]
mp3_google_search = ['MP3\here.mp3', 'MP3\search.mp3']
mp3_greeting_list = ['MP3\hope.mp3', 'MP3\how.mp3']
mp3_open_launch_list = ['MP3\open.mp3', 'MP3\launch.mp3']

error_occurrence = 0


def is_valid_google_search(phrase):
    if google_searches_dict.get(phrase.split(' ')[0]) == phrase.split(' ')[0]:
        return True


def is_valid_youtube_search(phrase):
    if youtube_search_dict.get(phrase.split(' ')[0] == phrase.split(' ')[0]):
        return True


def is_valid_youtube_song(phrase):
    if youtube_song_dict.get(phrase.split(' ')[0] == phrase.split(' ')[0]):
        return True

def time_():
    Time=datetime.datetime.now().strftime("%H:%M:%S")
    speak("The current time is")
    pyttsx3.speak(Time)

def play_sound(mp3_list):
    mp3 = random.choice(mp3_list)
    playsound(mp3)


def read_voice_cmd():
    voice_text = ''
    print('Listening...')

    global error_occurrence

    try:
        with sr.Microphone() as source:
            audio = speech.listen(source=source, timeout=10, phrase_time_limit=5)
        voice_text = speech.recognize_google(audio)
    except sr.UnknownValueError:

        if error_occurrence == 0:
            play_sound(mp3_listening_problem_list)
            error_occurrence += 1
        elif error_occurrence == 1:
            play_sound(mp3_struggling_list)
            error_occurrence += 1

    except sr.RequestError as e:
        print('Network error')
    except sr.WaitTimeoutError:

        if error_occurrence == 0:
            play_sound(mp3_listening_problem_list)
            error_occurrence += 1
        elif error_occurrence == 1:
            play_sound(mp3_struggling_list)
            error_occurrence += 1

    return voice_text


def is_valid_note(greet_dict, voice_note):
    for key, value in greet_dict.items():
        # 'Hello edith'
        try:

            if value == voice_note.split(' ')[0]:
                return True
                break
            elif key == voice_note.split(' ')[1]:
                return True
                break
        except IndexError:
            pass

    return False


def sendEmail(to, content):
    server=smtplib.SMTP('smtp.gmail.com', 587)
    server.ehlo()
    server.starttls()
    server.login('bagul.atharva009@gmail.com', 'fergussonkijai')
    server.sendmail('luciferpro36@gmail.com', to, content)
    server.close()
    pyttsx3.speak('Email sent!')


def screenshot():
    img = pyautogui.screenshot()
    img.save('C:\\Users\HP\Desktop\Friday\ss.png')


def cpu():
    usage = str(psutil.cpu_percent())
    pyttsx3.speak('CPU percentage is at' + usage)
    battery = psutil.sensors_battery()
    pyttsx3.speak("Battery percentage is at")
    pyttsx3.speak(battery.percent)


def jokes():
    pyttsx3.speak(pyjokes.get_joke())

if __name__ == '__main__':

    playsound('MP3\sahaya.mp3')

    while True:

        voice_note = read_voice_cmd().lower()
        print('cmd : {}'.format(voice_note))

        if is_valid_note(greeting_dict, voice_note):
            print('In greeting...')
            play_sound(mp3_greeting_list)
            continue
        elif is_valid_note(open_launch_dict, voice_note):
            print('In open...')
            play_sound(mp3_open_launch_list)
            if (is_valid_note(social_media_dict, voice_note)):
                # Launch Facebook
                key = voice_note.split(' ')[1]
                webbrowser.open(social_media_dict.get(key))
            else:
                os.system('explorer C:\\{}'.format(voice_note.replace('open ', '').replace('launch ', '')))
                print('explorer C:\\{}'.format(voice_note.replace('open ', '').replace('launch ', '')))

            continue
        elif 'time' in voice_note:
            time_()
            continue
        elif is_valid_google_search(voice_note):
            print("searching...")
            play_sound(mp3_google_search)
            webbrowser.open('https://www.google.co.in/search?q={}'.format(voice_note))
            continue
        elif 'thanks' in voice_note or 'thanks sahaya' in voice_note or 'thankyou' in voice_note or 'thankyou sahaya' in voice_note:
            print('Thanks boss...')
            play_sound(mp3_thankyou_list)
            continue
        elif 'wikipedia' in voice_note:
            speak("here is what i found on wikipedia...")
            query = voice_note.replace("wikipedia", "")
            result = wikipedia.summary(query, sentences=3)
            print(result)
            pyttsx3.speak(result)
            continue
        elif 'send email' in voice_note:
            try:
                pyttsx3.speak('What should i say?')
                content = voice_note()
                pyttsx3.speak("whom should i send?")
                receiver = input("Enter the receivers Email Address : ")
                to = receiver
                sendEmail(to, content)
                pyttsx3.speak("email has been sent !")
            except Exception as e:
                print(e)
                pyttsx3.speak('sorry sire! I am unable to send email at the moment, please try again later')
            continue
        elif 'tell me your name' in voice_note or 'your name' in voice_note:
            pyttsx3.speak("My friends call me sahaya")
            print("My friends call me sahaya")
            continue
        elif 'tell me your birth date' in voice_note:
            pyttsx3.speak("its twenty second of july twenty twenty ! Lahan aahe me ajun")
            print('lahan aahe mi ajun')
            continue
        elif 'news' in voice_note:
            try:
                jsonObj = urlopen("http://newsapi.org/v2/top-headlines?sources=TechCruncha&apiKey=4a70fa9c17414b7c87f09ff7e36371f2")
                data = json.load(jsonObj)
                i = 1

                pyttsx3.speak('here are some top news from the Tech Crunch')
                print('===========================NEWS========================='+'\n')
                for item in data['articles']:
                    print(str(i)+'. ' + item['title'] +'\n')
                    print(item['description']+'\n')
                    pyttsx3.speak(item['title'])
                    i += 1
            except Exception as e:
                print(str(e))
            continue
        elif 'locate' in voice_note:
            voice_note = voice_note.replace("locating", "")
            location = voice_note
            pyttsx3.speak("you asked to locate"+location)
            webbrowser.open_new_tab("http://www.google.com/maps/place/"+location)
            continue
        elif 'restart system' in voice_note or 'restart system sahaya' in voice_note:
            os.system("shutdown /r /t 1")
            continue
        elif 'shutdown system' in voice_note or 'shutdown system sahaya' in voice_note:
            os.system("shutdown /s /t 1")
            continue

        elif 'music' in voice_note or 'music sahaya' in voice_note:
            songs_dir = 'D:\\Music'
            songs = os.listdir(songs_dir)
            os.startfile(os.path.join(songs_dir, songs[0]))

            pyttsx3.speak('Okay sire here is your music! Enjoy!')
            continue
        elif 'make a note' in voice_note:
            pyttsx3.speak("What would you like me to write down  ?")
            data = read_voice_cmd()
            pyttsx3.speak("confirming what you said , you said : " + data)
            remember = open('data.txt', 'w')
            remember.write(data)
            remember.close()
            continue
        elif 'browse notes sahaya' in voice_note or 'browse notes' in voice_note:
            remember = open('data.txt', 'r')
            pyttsx3.speak("yes sire! there are a few.:" + remember.read())
            continue
        elif 'calculate' in voice_note.lower():
            app_id = "WOLFRAMALPHA_APP_ID"
            client = wolframalpha.Client('7JJ7JV-5TUPXPPXT6')
            index = voice_note.lower().split().index('calculate')
            query = voice_note.split()[index + 1:]
            res = client.query(' '.join(query))
            answer = next(res.results).text
            print(answer)
            pyttsx3.speak("The answer is " + answer)
            continue


        elif 'run diagnostics' in voice_note:
            cpu()
            continue
        elif ' tell me a joke  sahaya' in voice_note or 'joke' in voice_note:
            jokes()
            continue
        elif 'how are you' in voice_note or 'how are you sahaya' in voice_note:
            stMsgs = ['very good sire!', 'just doing my thing!', 'I am good and full of energy!']
            pyttsx3.speak(random.choice(stMsgs))
            continue
        elif 'define yourself' in voice_note or 'define yourself sahaya' in voice_note:
            pyttsx3.speak(
                "I am sahaaya,your personal assistant, a miracle developed in python! I am here to make your life easier. you can command me to perform various tasks")
            continue

        elif 'your developer' in voice_note:
            pyttsx3.speak("I was created by my god Master Atharva Baagul")
            print('I was created by my god Master Atharva Bagul')
            continue
        elif 'youtube' in voice_note or 'youtube sahaya' in voice_note:
            pyttsx3.speak("Opening youtube")
            webbrowser.open('https://www.youtube.com/search?q={}'.format(voice_note))
            continue
        elif 'song sahaya' in voice_note or 'song' in voice_note:
            pyttsx3.speak("Opening youtube music...enjoy!")
            webbrowser.open('https://music.youtube.com/search?q={}'.format(voice_note))
            continue
        elif 'jai hind dosto' in voice_note:
            pyttsx3.speak("jai hind , jai maharashtra")
            continue
        elif 'chrome sahaya' in voice_note or 'google chrome' in voice_note:
            pyttsx3.speak("opening chrome")
            os.startfile('C:\Program Files (x86)\Google\Chrome\Application\chrome.exe')
            continue
        elif 'word sahaya' in voice_note or 'msword' in voice_note:
            pyttsx3.speak("opening msword")
            os.startfile(
                'C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Microsoft Office\Microsoft Office Word 2007.lnk')
            continue
        elif 'powerpoint sahaya' in voice_note or 'mspowerpoint' in voice_note:
            pyttsx3.speak("opening mspowerpoint")
            os.startfile(
                'C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Microsoft Office\Microsoft Office PowerPoint 2007.lnk')
            continue
        elif 'excel sahaya' in voice_note or 'ms excel' in voice_note:
            pyttsx3.speak("opening MS excel")
            os.startfile(
                'C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Microsoft Office\Microsoft Office Excel 2007.lnk')
            continue
        elif 'camera' in voice_note or 'take a photo' in voice_note or 'photo sahaya' in voice_note:
            pyttsx3.speak("sure thing! smile please!")
            ec.capture(0, 'sahayaimg', "img.jpg")
            continue
        elif'dont listen' in voice_note or 'stop listening' in voice_note:
            pyttsx3.speak("oKay sir! for how much time you want me to stop  listening commands?")
            a = int(read_voice_cmd())
            time.sleep(a)
            print(a)
            continue
        elif 'd drive sahaya' in voice_note or ' D drive ' in voice_note:
            pyttsx3.speak("sure thing")
            os.startfile('D:\\')
        elif 'screenshot' in voice_note:
            pyttsx3.speak("screenshot taken")
            screenshot()
            continue
        elif 'command prompt' in voice_note or 'command prompt sahaya' in voice_note:
            pyttsx3.speak("Opening command prompt")
            os.startfile(
                'C:\\Users\HP\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\System Tools\Command Prompt.lnk')
            continue
        elif 'movies' in voice_note or 'movies sahaya ' in voice_note:
            pyttsx3.speak("opening Movies")
            os.startfile('D:\\movies')
            continue
        elif'control panel'in voice_note or 'control panel sahaya' in voice_note:
            pyttsx3.speak("opening control panel")
            os.startfile('C:\\Users\HP\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\System Tools\Control Panel.lnk')
            continue
        elif 'what are you upto' in voices or 'long time no see' in voice_note or 'good to see you' in voice_note:
            stMsgs=['yeah! was busy helping people', 'nivant challay jeevan', '']
            pyttsx3.speak(random.choice(stMsgs))
            continue
        elif 'how was your day' in voice_note or 'how are you doing' in voice_note:
            pyttsx3.speak("preety good sir, happy to help you")
            continue
        elif 'sleep' in voice_note:
            pyttsx3.speak('For how much time?')
            ans = int(voice_note)
            time.sleep(ans)
            print(ans)
        elif 'bye sahaya' in voice_note or 'goodbye sahaya' in voice_note or 'exit sahaya' in voice_note or 'jata' in voice_note or 'quit' in voice_note or 'go offline' in voice_note:
            print('Good bye Sire , have a nice day...')
            play_sound(mp3_bye)
            exit()
