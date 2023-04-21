import pyttsx3
import datetime
import speech_recognition as sr
import wikipedia
import webbrowser
import os
import random
import datetime
import GoogleNews
import feedparser
from win32com.client import Dispatch
import requests
import json
engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[0].id)
def speak(audio):
    engine.say(audio)
    engine.runAndWait()

def wishme():
    hour = int(datetime.datetime.now().hour)
    if hour>=0 and hour<=12:
        speak("Good morning Gaurav sir")
    elif hour>12 and hour<=18:
        speak("Good afternoon, Gaurav sir")
    else:
        speak("Good evening, Gaurav sir")

def getSpeech():
    r = sr.Recognizer()  
    with sr.Microphone() as source:
        print("Listening...")                                                                                       
        r.pause_threshold = 1
        r.energy_threshold = 200
        audio = r.listen(source) 

    try:
        print("Recognizing...")
        speak("Recognizing...")
        query = r.recognize_google(audio, language = 'en-in')   
        print(f"You said : {query}")
        

    except Exception as e:
        print(e)
        print("Can't hear")
        speak("Can't hear")
        return "None"
    return query        
                                           
if __name__ == "__main__":
    wishme()
    speak("iam your personal AI, how can i help you sir")
    while True:
        query = getSpeech().lower()
        if "you are ismart" in query:
            speak("Thankyou for your compliment, iam trying my best to become a good bot!")

        elif "who are you" in query:
            speak("My name is Ai, and iam a desktop assistant, iam programmed in python, im still in development and my developers is addidng more nerves in me, and aiming to become best bot!")
        
        elif "who created you" in query:
            speak("i was created by Aditya, an computer science engineer")

        elif "tell me about" in query:
            speak("ok, let me remember...")
            query = query.replace("tell me about", "")
            result = wikipedia.summary(query, sentences = 2)
            speak("Yeah, i got it, listen")
            speak(result)
            print(result)

        elif "what is" in query:
            speak("ok, let me remember...")
            query = query.replace("tell me about", "")
            result = wikipedia.summary(query, sentences = 2)
            speak("Yeah, i got it, listen")
            speak(result)
            print(result)

        elif "who is" in query:
            speak("ok, let me remember...")
            query = query.replace("tell me about", "")
            result = wikipedia.summary(query, sentences = 2)
            speak("Yeah, i got it, listen")
            speak(result)
            print(result)  

        elif "who was" in query:
            speak("ok, let me remember...")
            query = query.replace("tell me about", "")
            result = wikipedia.summary(query, sentences = 2)
            speak("Yeah, i got it, listen")
            speak(result)
            print(result)   

        elif "what was" in query:
            speak("ok, let me remember...")
            query = query.replace("tell me about", "")
            result = wikipedia.summary(query, sentences = 2)
            speak("Yeah, i got it, listen")
            speak(result)
            print(result)     

        elif "youtube" in query:
            try:
                speak("ok, opening youtube")
                print("opening YouTube...")
                webbrowser.open("youtube.com")
            except Exception as ex:
                print(ex)
                speak("sorry, i cant do that, there are some exceptions while execution your command") 
        
        elif "google" in query:
            try:
                speak("as you command, opening google")
                print("Opening Google...")
                webbrowser.open("google.co.in")
            except Exception as ex:
                print(ex)
                speak("sorry, i cant do that, there are some exceptions while execution your command")     

        elif "shutdown" in query:
            speak("welcome, standing-by for new orders")
            print("Re-run program to use again")
            break 
        
        elif "news" in query:
            speak("today's news feeds are,")
            try:
                url = "https://newsapi.org/v2/top-headlines?sources=bbc-news&apiKey=22fb4a7728be442fbaddfa826a35dad4"
                news = requests.get(url).text
                news_dict = json.loads(news)

                art = news_dict['articles']
                for articles in art:
                    speak(articles['title'])
            except Exception as ex:
                print(ex)
                speak("sorry, i cant do that, there are some exceptions while execution your command") 


        elif "send email" in query:
            speak("opening email...")
            print("opening email...")
            try:
                webbrowser.open("https://mail.google.com/mail/u/0/#inbox?compose=new")
            except Exception as ex:
                print(ex)
                speak("sorry, i cant do that, there are some exceptions while execution your command") 
        