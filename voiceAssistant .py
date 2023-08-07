from cProfile import label
from distutils import cmd
import subprocess
from tkinter.font import BOLD
import wolframalpha
import pyttsx3
import random
import speech_recognition as sr
import wikipedia
import webbrowser
import os
import winshell
# import pyjokes
import feedparser
import smtplib
import datetime 
import json
import requests
from twilio.rest import Client
from bs4 import BeautifulSoup
import win32com.client as wincl
from urllib.request import urlopen
import tkinter as tk
from tkinter import *
from tkinter import font as tkFont
import shutil



voiceEngine = pyttsx3.init('sapi5')
voices = voiceEngine.getProperty('voices')
voiceEngine.setProperty('voice', voices[1].id)


class VoiceAssistant(tk.Frame):
    def __init__(self,master) -> None:
        super().__init__(master)
        self.Label=tk.Label(wn, text='Welcome to Our Project: Personal Assistant', bg='black',fg='white', font=('Arial', 20))
        self.Label.place(x=100, y=30)

        self.Mic_icon = PhotoImage(file = r"D:\\VIT\\Sem 5\\SE Lab\\Assistant\\MIC.png")

        helv36 = tkFont.Font(family='Helvetica', size=16, weight='normal')
        tk.Button(wn, bg='black',image=self.Mic_icon,bd=0,state="normal",activebackground='black',
            command=self.callVoiceAssistant).place(x=280, y=130)

        Label(wn,text="Press to speak here",bg='black',fg='white', font=('Arial', 20)).place(x=300,y=500)
        Label(wn,text="Jessica : ",bg='black',fg='white', font=('Arial', 20)).place(x=30,y=600)
        self.speak_label=Label(wn, text=' ', bg='black',
            fg='white', font=('Courier', 20))
        self.speak_label.place(x=200, y=600)
        Label(wn,text="User : ",bg='black',fg='white', font=('Arial', 20)).place(x=30,y=700)
        self.comm_label=Label(wn, text=' ', bg='black',
            fg='white', font=('Courier', 20,BOLD))
        self.comm_label.place(x=200, y=700)
        wn.update_idletasks()
        b2 = Button(wn, text = "Exit",command = wn.destroy,bg='black',fg='white',bd=0,state="normal",activebackground="black",font=helv36)
        b2.place(x=800,y=30)

    def callVoiceAssistant(self):

        uname=''
        global asname
        os.system('cls')
        self.wish()
        self.getName()
        print(uname)

        while True:

            command = self.takeCommand()
            command.lower()
            print(command)
            if command != "None":
                if "Jessica" in command:
                    self.wish()
                    
                elif 'how are you' in command:
                    self.speak("I am fine, Thank you")
                    self.speak("How are you, ")
                    self.speak(uname)

                elif "good morning" in command or "good afternoon" in command or "good evening" in command:
                    self.speak("A very" +command)
                    self.speak("Thank you for wishing me! Hope you are doing well!")

                elif 'fine' in command or "good" in command:
                    self.speak("It's good to know that your fine")
            
                elif "who are you" in command:
                    self.speak("I am your virtual assistant.")

                elif "change my name" in command:
                    self.speak("What would you like me to call you, Sir or Madam ")
                    uname = self.takeCommand()
                    self.speak('Hello again,')
                    self.speak(uname)
                
                elif "change name" in command:
                    self.speak("What would you like to call me, Sir or Madam ")
                    asname = self.takeCommand()
                    self.speak("Thank you for naming me!")

                elif "what's your name" in command:
                    self.speak("People call me")
                    self.speak(asname)
                
                elif 'time' in command:
                    strTime = datetime.datetime.now()
                    curTime=str(strTime.hour)+"hours"+str(strTime.minute)+"minutes"+str(strTime.second)+"seconds"
                    self.speak(uname)
                    self.speak(f" the time is {curTime}")
                    print(curTime)

                elif 'wikipedia' in command:
                    self.speak('Searching Wikipedia')
                    command = command.replace("wikipedia", "")
                    results = wikipedia.summary(command, sentences = 3)
                    self.speak("These are the results from Wikipedia")
                    print(results)
                    self.speak(results)

                elif 'youtube' in command:
                    self.speak("Here you go, the Youtube is opening\n")
                    webbrowser.open("youtube.com")

                elif 'google' in command:
                    self.speak("Opening Google\n")
                    webbrowser.open("google.com")

                elif 'play music' in command or "play song" in command:
                    self.speak("Enjoy the music!")
                    music_dir = "C:\\Users\\basul\\Music\\Playlists"
                    songs = os.listdir(music_dir)
                    print(songs)
                    random = os.startfile(os.path.join(music_dir, songs[1]))

                    
                elif 'mail' in command:
                    try:
                        self.speak("Whom should I send the mail")
                        to = input()
                        self.speak("What is the body?")
                        content = self.takeCommand()
                        self.sendEmail(to, content)
                        self.speak("Email has been sent successfully !")
                    except Exception as e:
                        print(e)
                        self.speak("I am sorry, not able to send this email")

                elif  'exit' in command or  'close' in command or  'quit' in command :
                    self.speak("Closing")
                    exit()

                elif 'stop' in command  or 'thanks' in command or 'thank you' in command:
                    
                    self.speak("Thanks for giving me your time")
                    break
                    


                elif "my girlfriend" in command or "my gf" in command:
                    self.speak("I'm not sure about that, may be you should give me some time")

                elif "i love you" in command:
                    self.speak("Thank you! But, It's a pleasure to hear it from you.")

                elif "weather" in command:
                    self.speak(" Please tell your city name ")
                    print("City name : ")
                    cityName = self.takeCommand()
                    self.getWeather(cityName)

                elif "what is" in command or "who is" in command:
                    
                    client = wolframalpha.Client("A682U3-UE39885KWY")
                    res = client.query(command)
                    try:
                        print (next(res.results).text)
                        self.speak (next(res.results).text)
                    except StopIteration:
                        print ("No results")

                elif 'search' in command:
                    command = command.replace("search", " ")
                    webbrowser.open(command)

                elif 'news' in command:
                    self.getNews()
                
                elif "don't listen" in command or "stop listening" in command:
                    self.speak("for how much time you want to stop me from listening commands")
                    a = int(self.takeCommand())
                    datetime.time.sleep(a)
                    print(a)

                elif "camera" in command or "take a photo" in command:
                    ec.capture(0, "Jarvis Camera ", "img.jpg")
                
                elif 'shutdown system' in command:
                        self.speak("Hold On a Sec ! Your system is on its way to shut down")
                        subprocess.call('shutdown / p /f')

                elif "restart" in command:
                    subprocess.call(["shutdown", "/r"])

                elif "sleep" in command:
                    self.speak("Setting in sleep mode")
                    subprocess.call("shutdown / h")

                elif "write a note" in command:
                    self.speak("What should i write, sir")
                    note = self.takeCommand()
                    file = open('jarvis.txt', 'w')
                    self.speak("Sir, Should i include date and time")
                    snfm = self.takeCommand()
                    if 'yes' in snfm or 'sure' in snfm:
                        strTime = datetime.datetime.now().strftime("%H: %M: %S")
                        file.write(strTime)
                        file.write(" :- ")
                        file.write(note)
                    else:
                        file.write(note)

                else:
                    self.speak("Sorry, I am not able to understand you")
            else:
                continue

    def wish(self):
        print("Wishing.")
        time = int(datetime.datetime.now().hour)
        global uname,asname
        if time>= 0 and time<12:
            self.speak("Good Morning!")

        elif time<17:
            self.speak("Good Afternoon !")

        else:
            self.speak("Good Evening !")

        asname ="JESSICA"
        self.speak("I am your Voice Assistant "+asname)
        print("I am your Voice Assistant,",asname)
        
    def getName(self):
        global uname
        self.speak("Can I please know your name?")
        
        uname=self.takeCommand()
        print("Name:",uname)
        self.speak("I am glad to know you!")
        columns = shutil.get_terminal_size().columns
        self.speak("How can i Help you, "+uname)

    def takeCommand(self):
        recog = sr.Recognizer()
        
        with sr.Microphone() as source:
            
            print("Listening to the user")
            self.speak_label.config(text="Listening...")
            app.update()

            recog.pause_threshold = 1
            userInput = recog.listen(source)

        try:
            print("Recognizing the command")
            command = recog.recognize_google(userInput, language ='en-in')
            print(f"Command is: {command}\n")

        except Exception as e:
            print(e)
            print("Unable to Recognize the voice.")
            return "None"
        

        self.comm_label.config(text=str(command))
        app.update()
        return command


    def sendEmail(to, content):
        print("Sending mail to ", to)
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.ehlo()
        server.starttls()
        #paste your email id and password in the respective places
        server.login('your email id', 'password') 
        server.sendmail('your email id', to, content)
        server.close()

    def getWeather(self,city_name):
        cityName=city_name #getting input of name of the place from user
        baseUrl = "http://api.openweathermap.org/data/2.5/weather?" #base url from where we extract weather report
        url = baseUrl + "appid=" + 'd850f7f52bf19300a9eb4b0aa6b80f0d' + "&q=" + cityName  
        response = requests.get(url)
        x = response.json()

        #If there is no error, getting all the wather conditions
        if x["cod"] != "404":
            y = x["main"]
            temp = y["temp"]
            temp-=273
            temp=float(temp)
            temp=round(temp,2)
            pressure = y["pressure"]
            humidiy = y["humidity"]
            desc = x["weather"]
            description = desc[0]["description"]
            info=(" The Temperature in "+cityName+" is " +str(temp)+"°C"+"\n with atmospheric pressure "+str(pressure) +"\n and humidity of " +str(humidiy)+"%" +"\n and " +str(description))
            print(info)
            self.speak(info)
        else:
            self.speak(" City Not Found ")

    def getNews(self):
        try:
            response = requests.get('https://www.bbc.com/news')
    
            b4soup = BeautifulSoup(response.text, 'html.parser')
            headLines = b4soup.find('body').find_all('h3')
            unwantedLines = ['BBC World News TV', 'BBC World Service Radio',
                        'News daily newsletter', 'Mobile app', 'Get in touch']

            for x in list(dict.fromkeys(headLines))[0:4]:
                if x.text.strip() not in unwantedLines:
                    print(x.text.strip())
                    self.speak(x.text.strip())
        except Exception as e:
            print(str(e))
    def speak(self,text):
        self.speak_label.config(text=text)
        app.update()
        voiceEngine.say(text)
        voiceEngine.runAndWait()

#Creating the main window 
wn = tk.Tk() 
wn.title("Voice Assistant")
wn.geometry('900x800')
wn.config(bg='black')
  
app=VoiceAssistant(master=wn)
app.after(1000,lambda: app.focus_force())
#Runs the window till it is closed
app.mainloop()