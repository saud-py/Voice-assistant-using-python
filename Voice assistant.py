#importing module
import subprocess
import wolframalpha
import pyttsx3
import tkinter
import json
import random
import operator
import speech_recognition as sr
import datetime
import wikipedia
import webbrowser
import os
import winshell
import pyjokes
import feedparser
import ctypes
import time
import requests
import shutil
from clint.textui import progress
from bs4 import BeautifulSoup
import win32com.client as wincl
from urllib.request import urlopen

print('Initializing your virtual assistant')

#setting the voice properties
engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[1].id) #standard code

#speak function 
def speak(audio):
    engine.say(audio)
    engine.runAndWait()

#greeting function
def wishMe():
    hour = int(datetime.datetime.now().hour)
    if hour > 0 and hour < 12:
        speak("Good Morning sir !")
    elif hour >= 12 and hour < 4:
        speak("Good Afternoon sir !")
    else:
        speak('Good Evening sir !')

    global assname
    assname = ("Sam")
    speak('I am your assistant')
    speak(assname)

#Defining the username to the program
def username():
    speak("What should i call you sir?")
    username = takeCommand()
    speak("Welcome master")
    speak(username)
    columns = shutil.get_terminal_size().columns

    print("*-*-*-*-*-*-*".center(columns))
    print("Welcome Mr.", username.center(columns))
    print("*-*-*-*-*-*-*".center(columns))

    speak("How can i help you, sir?")


def takeCommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        r.adjust_for_ambient_noise(source, duration=0.2)
        print("Listening...")

        r.pause_threshold = 0.5
        audio = r.listen(source)

    try:
        print("recognizing...")
        query = r.recognize_google(audio, language='en-in')
        print(f"User said : {query}\n")

    except Exception as e:
        print(e)
        speak("Sorry, can you repeat that again?")
        return "None"
    return query


if __name__ == "__main__":
    clear = lambda: os.system('cls')
    clear()
    wishMe()
    username()

    while True:
        query = takeCommand().lower()
        if 'google' in query:
            speak("Searching in wikipedia")
            query = query.replace("google", "")
            results = wikipedia.summary(query, sentences=2)
            speak("According to wikipedia")
            print(results)
            speak(results)
        elif 'open youtube' in query:
            webbrowser.open("youtube.com")
            speak("Youtube is opened")
        elif 'open google' in query:
            webbrowser.open('google.com')
            speak("Google is opened")
        elif 'open gmail' in query:
            webbrowser.open("gmail.com")
            speak("Gmail is opened")
        elif 'the time' in query:
            strTime = datetime.datetime.now().strftime("%H:%M:%S")
            speak(f"Sir, the time is {strTime}")
        elif 'how are you' in query:
            speak('I am fine, Thank you')
            speak('How are you, sir?')
        elif 'fine' in query or "good" in query:
            speak('It is good to know that you are fine')
        elif "change my name to" in query:
            query = query.replace("change my name to", "")
            assname = query
        elif "change name" in query:
            speak("What would you like to call, sir")
            assname = takeCommand()
            speak("Thanks for naming me")
        elif "what is your name" in query:
            speak("My friends call me")
            speak(assname)
            print("My friends call me", assname)
        elif 'exit' in query:
            speak("Thanks for giving me your time")
            exit()
        elif "who made you" in query or "who created you" in query:
            speak("I was designed by Saud, developed in python envirnoment, I have many features and functions added yet i may not be user interactive")
        elif 'joke' in query:
            speak(pyjokes.get_jokes())
        elif "calculate" in query:
            app_id = "XJ6TT7-X9YXAUXR65"
            client = wolframalpha.Client(app_id)
            indx = query.lower().split.index('calculate')
            query = query.split()[indx + 1:]
            res = client.query(''.join(query))
            answer = next(res.results).text
            print("The answer is" + answer)
            print("The answer is" + answer)
        elif "who am I " in query:
            speak("As long as you speak you must definitely be a human")
        elif 'news' in query:
            try:
                jsonObj = urlopen(''https://newsapi.org / v1 / articles?source = the-times-of-india&sortBy = top&apiKey =\\times of India Api key\\'')
                data = json.load(jsonObj)
                i = 1
                speak('here are some of the top news from The times of India')
                print('''======== TIMES OF INDIA ========''')
                for item in data['articles']:
                    print(str(i) + '. ' + item['title'] + '\n')
                    print(item['description'] + '\n')
                    speak(str(i) + '. ' + item['title'] + '\n')
                    i += 1
            except Exception as e:

                print(str(e))

        elif 'lock window' in query:
            speak("locking the device")
            ctypes.windll.user32.LockWorkStation()
        elif 'shutdown system' in query:
            speak("Hold On a Sec ! Your system is on its way to shut down")
            subprocess.call('shutdown / p /f')
        elif 'empty recycle bin' in query:
            winshell.recycle_bin().empty(confirm=False, show_progress=False, sound=True)
            speak("Recycle Bin Recycled")
        elif "don't listen" in query or "stop listening" in query:
            speak("for how much time you want to stop me from listening commands")
            a = int(takeCommand())
            time.sleep(a)
            print(a)
        elif "where is" in query:
            query = query.replace("where is", "")
            location = query
            speak("User asked to Locate")
            speak(location)
            webbrowser.open("https://www.google.com/maps/place/" + location + "")
        elif "write a note" in query:
            speak("What should i write, sir")
            note = takeCommand()
            file = open('jarvis.txt', 'w')
            speak("Sir, Should i include date and time")
            snfm = takeCommand()
            if 'yes' in snfm or 'sure' in snfm:
                strTime = datetime.datetime.now().strftime("%H:%M:%S")
                file.write(strTime)
                file.write(" :- ")
                file.write(note)
            else:
                file.write(note)
        elif "show note" in query:
            speak("Showing Notes")
            file = open("jarvis.txt", "r")
            print(file.read())
            speak(file.read(6))
        elif "jarvis" in query:

            wishMe()
            speak("Jarvis 1 point o in your service Mister")
            speak(assname)
        elif "weather" in query:

            # Google Open weather website
            # to get API of Open weather
            api_key = "faec2c00df5b6368e7e89f3d8673bda8"
            base_url = "https://api.meteomatics.com/2021-12-06T00:00:00Z/t_2m:C/52.520551,13.461804/html"
            speak(" City name ")
            print("City name : ")
            city_name = takeCommand()
            complete_url = base_url + "appid =" + api_key + "&q =" + city_name
            response = requests.get(complete_url)
            x = response.json()

            if x["cod"] != "404":
                y = x["main"]
                current_temperature = y["temp"]
                current_pressure = y["pressure"]
                current_humidiy = y["humidity"]
                z = x["weather"]
                weather_description = z[0]["description"]
                print(" Temperature (in kelvin unit) = " + str(
                    current_temperature) + "\n atmospheric pressure (in hPa unit) =" + str(
                    current_pressure) + "\n humidity (in percentage) = " + str(
                    current_humidiy) + "\n description = " + str(weather_description))
            else:
                speak("City not found")
        elif "Good morning" in query:
            speak("A very warm" + query)
            speak("Have a nice day at work today, and do check for the weather")
        elif "what is " in query or "who is" in query:
            client = wolframalpha.Client("XJ6TT7-X9YXAUXR65")
            res = client.query(query)
        try:
            print(next(res.results).text)
            speak(next(res.results).text)
        except StopIteration:
            print("No results")
