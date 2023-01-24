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
import smtplib
import ctypes
import time
import requests
import shutil
#from twilio.rest import client
from clint.textui import progress
from ecapture import ecapture as ec
from bs4 import BeautifulSoup
import win32com.client as wincl
from urllib.request import urlopen

# for text and speech will set our engine to Pyttsx3
#sapi5 is a MS speech application platform interface which will be used for text to speech function
engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[1].id) #1 is for female voice and 0 for male voice

def speak(audio):
    engine.say(audio)
    engine.runAndWait()

def wishMe():
    hour = int(datetime.datetime.now().hour)
    if hour>= 0 and hour<12:
        speak("Good Morning Madam !")

    elif hour>=12 and hour<18:
        speak("Good Afternoon Madam !")

    else:
        speak("Good Evening Madam !")

    assname =("Penny")
    speak("I am your Assistant")
    speak(assname)

def username():
    speak("What should I call you Madam")
    uname = takeCommand()
    speak("Welcome Miss")
    speak(uname)
    columns = shutil.get_terminal_size().columns

    print("#####################".center(columns))
    print("Welcome Miss. ", uname.center(columns))

    speak("How can I help you, Madam")


def takeCommand():
    r = sr.Recognizer()

    with sr.Microphone() as source:

        print("Listening...")
        r.pause_threshold = 1
        audio = r.listen(source)

    try:
        print("Recognizing...")
        query = r.recognize_google(audio, language = 'en-in')
        print(f"User said: {query}]n")

    except Exception as e:
        print(e)
        print("Unable to Recognize your voice")
        return "None"
    return query

#what does this mean??
    def sendEmail(to, content):
        server =smtplib.SMTP('smtp.gmail.com', 587) 
        server.ehlo()
        server.starttls()

        #Enable low security in gmail
        server.login('your email id', 'your email password')
        server.sendmail('your email id', to, content)
        server.close()


if __name__ == '__main__':
    clear = lambda: os.system('cls') #this function will clean any command before executing this python file

    clear()
    wishMe()
    username()

    while True:
        query = takeCommand().lower() #all user commands will be stored in query in lower case for easy recognition

        if 'wikipedia' in query:
            speak('Searching Wikipedia...')
            query = query.replace("wikipedia," "")
            results = wikipedia.summary(query, sentences = 3)
            speak("According to Wikipedia")
            print(results)
            speak(results)

        elif 'open youtube' in query:
            speak("Here you go to Youtube\n")
            webbrowser.open("youtube.com")

        elif 'open google' in query:
            speak("Here you go to Google\n")
            webbrowser.open("google.com")

        elif 'stakeoverflow' in query:
            speak("Here you go to StackOverflow. Happy coding!!")
            webbrowser.open("stackoverflow.com")

        # elif 'play music'in query or "play song" in query:
        #     speak("Here you go with music")
        #     #music_dir = "G:\\Song"
        #     music_dir = "C:\Users\laven\Music\Music"
        #     songs = os.listdir(music_dir)
        #     print(songs)
        #     random = os.startfile(os.path.join(music_dir, songs[1]))

        elif 'the time' in query:
            strTime = datetime.datetime.now().strftime("% H:% M:% S")
            speak(f"Madam, the time is {strTime}")

        elif 'email to lavender' in query:
            try:
                speak("What should I say?")
                content = takeCommand()
                to = "Rceiver email address"
                sendEmail(to, content)
                speak("Email has been sent !")
            except Exception as e:
                print(e)
                speak("I am not able to send this email")

        elif 'send a mail' in query:
            try:
                speak("What should I say?")
                content = takeCommand()
                speak("whom should I send")
                to = input()
                sendEmail(to, content)
                speak("Email has been sent !")
            except Exception as e:
                print(e)
                speak("I am not able to send this email")

        elif 'how are you' in query:
            speak("I am fine, Thank you")
            speak("How are you, Madam")

        elif 'fine' in query or "good" in query:
            speak("It's good to know that you are fine")

        elif "change my name to" in query:
            query = query.replace("change my name to", "")
            assname = query

        elif "change my name " in query:
            speak("What would you like to call me, Madam")
            assname = takeCommand()
            speak("Thanks for naming me")

        elif "what's your name" in query or "what is your name" in query:
            speak("My friends call me", assname)
            speak(assname)
            print("My friends call me", assname)

        elif 'exit' in query:
            speak("Thanks for giving me your time")
            exit()

        elif "who made you" in query or "who created you" in query:
            speak("I have been created by Lavender.")

        elif 'joke' in query:
            speak(pyjokes.get_joke())

        elif "calculate" in query:

            app_id = "wolframalpha api id"
            client = wolframalpha.Client(app_id)
            indx = query.lower().split().index('calculate')
            query = query.split()[indx + 1:]
            res = client.query(''.join(query))
            answer = next(res.results).text
            print("The answer is" + answer)
            speak("The answer is" + answer)

        elif 'search' in query or 'play' in query:

            query = query.replace("search", "")
            query = query.replace("play", "")
            webbrowser.open(query)

        elif "who am I" in query:
            speak("if you talk then definitely you are human")

        elif "why you came to the world" in query:
            speak("Thanks to Lavender. Further it's a secret")

        elif 'power point presentation' in query:
            speak("opening Power point presentation")
            power = r":\Users\laven\Documents\Network security\cse Group 2.pptx"

            
        elif 'is love' in query:
            speak("It is 7th sense that destroy all other senses")
 
        elif "who are you" in query:
            speak("I am your virtual assistant created by Lavender")
 
        elif 'reason for you' in query:
            speak("I was created as a Minor project by Miss Lavender ")

        elif 'change background' in query:
            ctypes.windll.user32.SystemParametersInfoW(20, 0,"Location of wallpaper",  0)
            speak("Background changed successfully")

        elif 'lock window' in query:
            speak("locking device")
            ctypes.windll.user32.LockworkStation()

        elif 'shutdown system' in query:
            speak("After downloading file please replace this file with the downloaded one")
            subprocess.call('shutdown / p / f')

        elif 'empty recycle bin' in query:
            winshell.recycle_bin().empty(confirm = False, show_progress = False, sound = True)
            speak("Recycle Bin recycled")

        elif "don't listen" in query or "stop listening" in query:
            speak("for how much time you want to stop Penny from listening commands")
            a = int(takeCommand())
            time.sleep(a)
            print(a)

        elif "where is" in query:
            query = query.replace("where is", "")
            location = query
            speak("User asked to Locate")
            speak(location)
            webbrowser.open("https://www.google.nl / maps / place/" + location + "")

        elif "camera" in query or "take a photo" in query:
            ec.capture(0, "Penny Camera", "img.jpg")

        elif "restart" in query:
            subprocess.call(["shutdown", "/r"])

        elif "hibernate" in query or "sleep" in query:
            speak("Hibernating")
            subprocess.call("shutdown / h")

        elif "log off" in query or "sign out" in query:
            speak("MAke sure all applications are closed before signing out")
            time.sleep(5)
            subprocess.call(["shutdown", "/1"])

        elif "write a note" in query:
            speak("What should I write, Madam")
            note = takeCommand()
            file = open('Penny.txt', 'w')
            speak("Madam, should i include date and time")
            snfm = takeCommand()
            if 'yes'in snfm or 'sure' in snfm:
                strTime = datetime.datetime.now().strftime("% H:% M:% S")
                file.write(strTime)
                file.write(" :- ")
                file.write(note)
            else:
                file.write(note)
                
        elif "show note" in query:
            speak("showing notes")
            file = open("Penny.txt", "r")
            print(file.read())
            speak(file.read(6))
#what does the following mean or do
        elif "update assistant" in query:
            speak("After downloading file please replace thisfile with the with the downloaded one")
            url = '# url after uploading file'
            r = requests.get(url, stream = True)

            with open("Voice.py", "wb") as Pydf: #what is this?

                total_length = int(r.headers.get('content-length'))

                for ch in progress.bar(r.iter_content(chunk_size= 2391975), expected_size =(total_length/total_length/1024 )+1):

                    if ch:
                        Pydf.write(ch)

        # NPPR9-FWDCX-D2C8J-H872K-2YT43
        elif "Penny" in query:

            wishMe()
            speak("Penny 1 point o in you service Mister")
            speak(assname)

        elif "weather" in query:
            # Google Open weather website
            # to get API of Open weather
            api_key = "Api key"
            base_url = "http://api.openweathermap.org/data/2.5/weather?"
            speak("City name")
            print("City name : ")
            city_name = takeCommand()
            complete_url = base_url + "appid =" + api_key + "&q =" + city_name
            response = requests.get(complete_url)
            x = response.json()

            if x["code"] != "404":
                y = x["main"]
                current_temperature = y["temp"]
                current_pressure = y["pressure"]
                current_humidity = y["humudity"]
                z = x["weather"]
                weather_description = z[0]["description"]
                print(" Temperature (in kelvin unit) = " +str(current_temperature)+"\n atmospheric pressure (in hPa unit) ="+str(current_pressure) +"\n humidity (in percentage) = " +str(current_humidity) +"\n description = " +str(weather_description))

            else:
                speak("City Not Found")
        
        # elif "send message" in query:
        #     #create an account on twilio to use this service
        #     account_sid = 'Account Sid key'
        #     auth_token = 'Auth token'
        #     client = client(account_sid, auth_token)

        #     message = client.message \
        #                     .create(
        #                         body = takeCommand(),
        #                         from_ = 'Sender No',
        #                         to = 'Receiver No'
        #                     )
        #     print(message.sid)

        elif "wikipedia" in query:
            webbrowser.open("wikipedia.com")

        elif "Good Morning" in query:
            speak("A warm" +query)
            speak("How are you Miss")
            speak(assname)

        # most asked question from google Assistant
        elif "will you be my girlfriend" in query or "will you be my boyfriend" in query:
            speak("I'm not sure about that. maybe you should give me sometime")

        elif "how are you" in query:
            speak("I'm fine, thanks for asking")

        elif "i love you" in query:
            speak("i love you too")

        elif "what is" in query or "who is" in query:

            #use the api we generated

            client =wolframalpha("API_ID")
            res = client.query(query)

            try:
                print(next(res.results).text)
                speak(next(res.results).text)
            except StopIteration:
                print ("No results")



        

