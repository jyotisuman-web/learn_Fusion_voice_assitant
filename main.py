import subprocess
import wolframalpha
import pyttsx3
import json
import Pyaudio
import speech_recognition as sr
import datetime
import wikipedia
import webbrowser
import os
import pyjokes
import requests
import shutil
import winshell
import sys
from clint.textui import progress
from bs4 import BeautifulSoup
import win32com.client as wincl
from urllib.request import urlopen
from bs4 import BeautifulSoup
engine=pyttsx3.init('sapi5')
voices=engine.getProperty('voices')
engine.setProperty('voice',voices[1].id)
#from pyttsx3
def speak(audio):
    engine.say(audio)
    engine.runAndWait()
def WishMe():
    hour=int(datetime.datetime.now().hour)
    if hour>=0 and hour<12:
        speak("Good Morning!")
        print("Good Morning!")

       
    elif hour>=12 and hour<18:
        speak("Good Afternoon!")
        print("Good Afternoon!")

    else:
        speak("Good Evening")
        print("Good Evening")

    assisname="Toki"
    speak("Hello, I am your assistant.")
    print("Hello, I am your assistant.")

    speak(assisname)
def username():
    speak("Please tell me your name.")
    print("Please tell me your name.")
    uname=takecommand()
    speak("Welcome")
    speak(uname)
    columns=shutil.get_terminal_size().columns
    print("*****************".center(columns))
    print("Welcome", uname.center(columns))
    print("*****************".center(columns))
    speak("Please tell me, how can I help you.")
def takecommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening...")
        r.pause_threshold = 1
        audio = r.listen(source)
    try:
        print("Recognizing...")
        query = r.recognize_google(audio, language='en-in')
        print(f"User said: {query}\n")
    except Exception as e:
        print(e)
        print("Unable to Recognize your voice. Please speak again.")
        return "None"
    return query



#Main Function
if __name__=='__main__':
    clear=lambda: os.system('cls')
    clear()
    WishMe()
    username()
    while True:
        query=takecommand().lower()
        if 'wikipedia' in query:
            speak('Searcing Wikipedia')
            query=query.replace('wikipedia',"")
            results=wikipedia.summary(query,sentences=3)
            speak("According to Wikipedia")
            print(results)
            speak(results)
        elif 'open youtube' in query:
            speak('Opening youtube\n')
            webbrowser.open("youtube.com")
        elif 'open google' in query:
            speak('Opening google\n')
            webbrowser.open("google.com")
        elif 'play music' in query or 'play song' in query:
            speak("Now listen the awesome music")
            music_dir="Music"
            songs=os.listdir(music_dir)
            print(songs)
            random_song=os.startfile(os.path.join(music_dir,songs[0]))
        elif 'open crome' in query:
            speak("Opening Crome")
            webbrowser.open("crome.com")
        elif 'the time' in query:
            strTime = datetime.datetime.now().strftime("% H:% M:% S")
            speak(f"the time is {strTime}")
        
        elif 'exit' in query or 'bye' in query:
            speak("Thank you, Come again.")
            sys.exit()
        elif 'sleep' in query:
            speak("Going to sleep")
            subprocess.call("shutdown / h")
        elif 'joke' in query:
            speak(pyjokes.get_joke())
        elif "calculate" in query or 'find' in query:
            app_id="PPR269-89Y6VQP44E"
            client = wolframalpha.Client(app_id)
            indx=query.lower().split().index('calculate')
            query=query.split()[indx +1:]
            res=client.query(' '.join(query))
            answer=next(res.results).text
            print("The answer is " + answer)
            speak("The answer is " + answer)
        elif 'Document' in query:
            speak("Opening Documents")
            power=r"C:\Users\HP\OneDrive - Rajasthan Technical University\Documents"
            os.startfile(power)
        elif 'news' in query:
            try:
                jsonObj=urlopen('https;://newsapi.org/vi/articles?source=the-time-of-india&sortBy=top&apiKey=\\times of India Api Key\\')
                data=json.load(jsonObj)
                i=1
                speak("today's news is")
                print('''*****Times of India*****''' + '\n')
                for item in data['articles']:
                    print(str(i) + '.'+ item['title'] + '\n')
                    print(item['description'] + '\n')
                    speak(str[i] + '.' + item['title'] + '\n')
                    i+=1
            except Exception as e:
                print(str(e))
        elif 'shutdown system' in query:
            speak('wait for a while! your system is shutting down')
            subprocess.call('shutdown / p / f')
        elif 'empty recycle bin' in query:
            winshell.recycle_bin().empty(confirm=False,show_progress=False,sound=True)
            speak("Recycle Bin Recycled")
        
        elif 'restart' in query: 
            subprocess.call(["shutdown","/r"])
        elif "write a note" in query:
            speak('What should I write!')
            note=takecommand()
            file=open('toki.txt','w')
            speak('should i include date and time to it.')
            sld=takecommand()
            if 'yes' in sld or 'sure' in sld or 'ok' in sld:
                strtime=datetime.datetime.now().strftime("%H:%M:%S")
                file.write(strtime)
                file.write(" :  ")
                file.write(note)
            else:
                file.write(note)
        elif 'show note' in query:
            speak("Here, is the note")
            file=open("toki.txt","r")
            print(file.read(6))
        elif 'toki' in query:
            WishMe()
            speak("Yes, you need help")
        elif "weather" in query:
            api_key="9f7bd6b5543662153f6a5a579b3cd79b"      
            url="http://api.openweathermap.org/data/2.5/weather?"
            speak("city name")
            print("City name: ")
            city_name=takecommand()
            complete_url=url+"appid=" + api_key + "&q=" + city_name 
            response=requests.get(complete_url)
            x=response.json()
            if x["code"]!="404":
                y=x["main"]
                current_temp=y["temp"]
                current_humidity=y["humidity"]
                z=x["weather"]
                weather_description=z[0]["description"]
                print("Temprature(in kelvin) = "+str(current_temp)+"\n Humidity(in percentage) = "+
                      str(current_humidity)+"\n Description = "+str(weather_description))
            else:
                speak("City not found")

        elif "what is" in query or "who is" in query:
            client=wolframalpha.Client("PPR269-89Y6VQP44E")
            res=client.query(query)
            try:
                print(next(res.results).text)
                speak(next(res.results).text)
            except StopIteration:
                print("No result")






     






          
