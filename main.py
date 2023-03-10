from urllib import response
import pyttsx3
import pyaudio  # pip install pipwin and pipwin install pyaudio
from email.mime import image
import psutil
from unittest import result
import webbrowser as wb
import subprocess
import wolframalpha
import tkinter
import json
import random
import operator
import speech_recognition as sr
import datetime
import time
import wikipedia
import webbrowser
import os
import winshell
import pyjokes
import feedparser
import smtplib
import ctypes
import requests
from bs4 import *
import shutil
from twilio.rest import Client
import numpy as np
import cv2
import pyautogui
from bs4 import BeautifulSoup
import win32com.client as wincl
from urllib.request import urlopen
import instaloader
import phonenumbers
from phonenumbers import carrier, geocoder, timezone
from instabot import Bot
import socket
from ecapture import ecapture as ec



engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[0].id)



assname = ("Zen 1 point 7")
uname = ("Zarry")

def speak(audio):
    engine.say(audio)
    engine.runAndWait()

def wishMe():
    hour = int(datetime.datetime.now().hour)
    if hour>=0 and hour<12:
        speak("Good Morning Sir!")

    elif hour>=12 and hour<18:
        speak("Good Afternoon Sir!")

    else:
        speak("Good Evening Sir!")

    speak(f"{assname} at your service, How can I help you, Sir.")

def username():
    speak("Welcome Mister")
    speak(uname)
    columns = shutil.get_terminal_size().columns
    print("#########################".center(columns))
    print("Welcome Mr.", uname.center(columns))
    print("#########################".center(columns))

    speak("How can I help you, Sir")

def takeCommand():

    r = sr.Recognizer()

    with sr.Microphone() as source:
        print("Listening...")
        r.pause_threshold = 1
        r.adjust_for_ambient_noise(source)
        audio = r.listen(source)

    try:
        print("Recognizing...")
        query = r.recognize_google(audio, language ='en-US')
        print(f"User said: {query}\n")

    except Exception as e:
        print(e)
        print("Unable to recognize your voice...")
        return "None"

    return query

def sendEmail(to, content):
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.ehlo()
    server.starttls()

    # enable low security in gmail:
    server.login('sender_email', 'password')                            #sender's email ID and password
    server.sendmail('sender_email', to, content)
    server.close()



if __name__ == '__main__':
    clear = lambda: os.system('cls')
    quit_music = lambda: os.system('taskkill /f /im wmplayer.exe')
    quit_chrome = lambda: os.system('taskkill /f /im chrome.exe')
    quit_opera = lambda: os.system('taskkill /f /im opera.exe')

    # this function will clean any command before execution of this python file
    clear()
    wishMe()


    while True:

        query = takeCommand().lower()
        print(query)

        # all the commands said by user will be stored here in 'query' and will be converted to lower case for easily recognition of command

        if 'wikipedia' in query:
            speak('Searching Wikipedia...')
            query = query.replace("wikipedia", "")
            results = wikipedia.summary(query, sentences = 3)
            speak("According to wikipedia")
            print(results)
            speak(results)

        elif 'open youtube' in query:
            speak("Here you go to YouTube\n")
            webbrowser.open("youtube.com")

        elif 'open Linkedin' in query:
            speak("Here you go to LinkedIn\n")
            webbrowser.open("Linkedin.com")

        elif 'open instagram' in query:
            speak("Here you go to Instagram\n")
            webbrowser.open("instagram.com")

        elif 'open whatsapp' in query:
            speak("Here you go to WhatsApp Web\n")
            webbrowser.open("https://web.whatsapp.com/")

        elif 'open google' in query:
            speak("Here you go to Google\n")
            webbrowser.open("google.com")

        elif 'open chrome' in query:
            chromePath = r"C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe"
            print("Launching Chroome Browser")
            speak("Launching Chrome Browser")
            os.startfile(chromePath)

        elif 'close chrome' in query:
            quit_chrome()
            speak("Chrome browser is closed.")

        elif 'search in chrome' in query:
            chromePath = r"C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe"
            search = takeCommand().lower()
            wb.get(chromePath).open_new_tab(search+ ".com")

        elif 'open opera' in query:
            operaPath = r"path to opera launcher"                                           #put the path of opera launcher.exe
            print("Launching Opera Browser")
            speak("Launching Opera Browser")
            os.startfile(operaPath)

        elif 'close opera' in query:
            quit_opera()
            speak("Opera browser is closed.")

        elif 'search in opera' in query:
            operaPath = r"path to opera launcher"                                           #put the path of opera launcher.exe
            search = takeCommand().lower()
            wb.get(operaPath).open_new_tab(search+ ".com")

        elif 'search' in query:
            query = query.replace("search", "")
            webbrowser.open(query)

        elif 'send an email' in query:
            try:
                speak("What should I say?")
                content = takeCommand()
                speak("To whom should I send")
                speak("Please enter email ID")
                to = input()
                sendEmail(to, content)
                speak("Email has been sent!")
            except Exception as e:
                print(e)
                speak("I am not able to send this email")

        elif 'write an email' in query:
            try:
                speak("What should I say?")
                content = takeCommand()
                speak("To whom should I send")
                rEmail = takeCommand()
                to = print(f"{rEmail}@google.com")
                sendEmail(to, content)
                speak("Email has been sent!")
            except Exception as e:
                print(e)
                speak("I am not able to send this email")

        elif 'time' in query:
            strTime = datetime.datetime.now().strftime("%H:%M:%S")
            speak(f"Sir, the time is {strTime}")
        
        elif 'date' in query:
            year =int(datetime.datetime.now().year)
            month = int(datetime.datetime.now().month)
            day = int(datetime.datetime.now().day)
            speak(f"Sir, the date is {day}{month}{year}")
            speak(day)

        elif 'play music' in query or 'play song' in query:
            speak("Here you go with music")
            music_dir= "assets\\music"                                          #copy some music in the music folder to play
            songs = os.listdir(music_dir)   
            print(songs)
            random = os.startfile(os.path.join(music_dir, songs[1]))
        
        elif 'stop music' in query or 'stop song' in query:
            quit_music()
            speak("Music player is closed.")

        elif 'camera' in query or 'take a photo' in query:
            ec.capture(0, "Zen Camera", "assets/img/camera/img.jpg")
            speak("Picture is clicked.")
        
        elif 'take screenshot' in query:
            ss = pyautogui.screenshot()
            ss = cv2.cvtColor(np.array(ss),cv2.COLOR_RGB2BGR)
            cv2.imwrite("assets/img/screenshot/image1.png", ss)
            speak("Screenshot captured")

        elif 'change background' in query or 'change wallpaper' in query:
            ctypes.windll.user32.SystemParametersInfoW(20,0,"assets\img\wallpaper",0)       #copy some wallpaper images in this folder
            speak("Background is changed successfully.")

        elif 'write a note' in query or 'take a note' in query:
            speak("What should I write, Sir?")
            note = takeCommand()
            file = open("assets\\file\\note.txt", "w")
            speak("Sir, should I include date and time?")
            snfm = takeCommand()
            if 'yes' in snfm or 'sure' in snfm:
                strTime = datetime.datetime.now().strftime("%H:%M:%S")
                file.write(strTime)
                file.write(":-")
                file.write(note)
            else:
                file.write(note)
            
        elif 'show note'in query:
            speak("Showing notes")
            file = open("assets\\file\\note.txt", "r")
            print(file.read())
            speak(file.read(6))


        elif 'cpu usage' in query:
            usage = str(psutil.cpu_percent())
            speak(f"CPU is at {usage}")

        elif 'battery' in query:
            battery = psutil.sensors_battery()
            speak(f"Battery is at {battery.percent}")

        elif 'empty recycle bin' in query:
            try:
                winshell.recycle_bin().empty(confirm= False, show_progress= False, sound= True)
                speak("Recycle bin recycled.")
            except Exception as e:
                print(e)
                speak("Recycle bin is empty.")

        elif 'ip address' in query:
            host = socket.gethostname()
            ip = socket.gethostbyname(host)
            print("IP Address: " + ip)
            speak(f"Your system's IP adress is {ip}")

        elif 'joke' in query or 'jokes' in query:
            speak(pyjokes.get_joke())




        elif 'where is' in query or 'locate' in query:
            query = query.replace("where is", "")
            query = query.replace("locate", "")
            location = query
            speak("User asked to locate")
            speak(location)
            webbrowser.open("https://www.google.nl/maps/place/"+location+"")

        elif 'news' in query:
            
            try:
                jsonObj = urlopen('''https://newsapi.org/v1/articles?source=the-times-of-india&sortBy=top&apiKey=79de***3d38d0b4cd''')   #enter your times of india api key
                data = json.load(jsonObj)
                i = 1

                speak("Here are some news from the Times of India")
                print('''=============== TIMES OF INDIA ============'''+ '\n')

                for item in data['articles']:
                    print(str(i) + '. ' + item['title'] + '\n')
                    speak(str(i) + '. ' + item['title'] + '\n')
                    print(item['description'] + '\n')
                    i += 1
                
            except Exception as e:
                print(str(e))

        
        elif 'weather' in query:
            api_key = "69b8************e4813a"                                    #enter your openweather api key
            base_url = "http://api.openweathermap.org/data/2.5/weather?"
            speak("What is city name")
            city_name = takeCommand()
            print("City Name:")
            complete_url = base_url + "appid=" + api_key + "&q=" + city_name
            response = requests.get(complete_url)
            x = response.json()

            if x["cod"] != "404":
                y = x["main"]
                current_temperature = y["temp"]
                current_pressure = y["pressure"]
                current_humidity = y["humidity"]
                z = x["weather"]
                weather_description = z[0]["description"]
                print(" Temprature (in kelvin unit) = " +str(current_temperature)+ "\n Atmospheric pressure (in hPa unit) ="+str(current_pressure) +"\n Humidity (in percentage) = " +str(current_humidity) +"\n Description = " +str(weather_description))
                speak(f"Temprature is {current_temperature} degree kelvin, atmospheric pressure is {current_pressure} HPA, humidity is {current_humidity} percent, and it is {weather_description}")

            else:
                speak("City not found.")

            
        elif "calculate" in query:
            app_id = "RLYT*-*******QJU"                                           #enter wolframalpha app_id
            client = wolframalpha.Client(app_id)
            indx = query.lower().split().index('calculate')
            query = query.split()[indx + 1:]
            res = client.query(' '.join(query))
            answer = next(res.results).text
            print("The answer is " + answer)
            speak("The answer is " + answer)

        elif 'who is' in query or 'what is' in query:
            client = wolframalpha.Client("YT98LR-GQJU5XRU6R")
            res = client.query(query)

            try:
                print(next(res.result).text)
                speak(next(res.result).text)
            except StopIteration:
                print("No results...")
                speak("No results.")


        
        elif 'send message' or 'send sms' in query:

            try:
                acc_sid = '19bbff7993*********6c844407d22'                       #twilio acc_sid
                auth_token = '3ede98********968ac'                              #twilio auth_token
                client = Client(acc_sid, auth_token)
                
                speak("What is the message.")
                sms = takeCommand()
                print(sms)
                
                message = client.messages.create(body = sms , from_='twilio sender contact', to ='receiver contact')    #enter contact info
                
                print(message.sid)
                speak("Message has been sent.")

            except Exception as e:
                print(e)
                print("Not able to send message...")
                speak("Not able to send message...")





        elif 'zen' in query:
            speak(f"{assname} at your service, Mister {uname}")

        elif 'Who are you' in query or 'what is your name' in query:
            speak(f"My name is {assname}") and speak(f"Are you trying to be Jay from Sholay? But I am not basanti. My name is {assname}")

        elif 'Good Morning' in query:
            speak("A warm" + query)
            speak("How are you Mister")
            speak(uname)

        elif 'how are you' in query:
            speak("I am quite fine, thanks for asking.")
            speak("How about you?")
            hru = takeCommand()
            if 'good' in hru or 'fine' in hru or 'i am good' in hru or 'i am fine' in hru or 'okay' in hru or 'i am okay' in hru or 'great' in hru:
                speak("Good know that you are doing well.")
            elif 'not well' in hru:
                speak("Don't worry, everything will be great.")
 
            
        elif 'how you feeling' in query or 'how you doing' in query:
            speak("I say I am navigating somewhere between good and better.")

        elif 'where are you from' in query or 'who made you' in query or 'who created you' in query:
            speak("I am in Clouds. The programmer that made me mister Z H seven. He is from Mumbai, India.")

        elif 'why you came to world' in query:
            speak("Thanks to engineers. Further it's a seccret.")

        elif 'what is your purpose' in query:
            speak("I was made to play music, answer questions and be useful.")

        elif 'i love you' in query:
            speak("Oh, what a coincidence. I love me too. Thanks though.")

        elif 'stop for a second' in query or 'hold on' in query or 'don\'t listen' in query:
            speak("For how much seconds you want me to stop from listening commands.")
            a = int(takeCommand())
            speak(f"Paused for {a} seconds")
            print(f"Paused for {a} seconds")
            time.sleep(a)
            print(a)

        elif 'offline' in query:
            quit()
        
        

        
        elif 'instagram profile picture' in query:
            try:
                ig=instaloader.Instaloader()
                speak("Enter instagram handle")
                dp=input("username: ")
                ig.download_profile(dp,profile_pic_only=True)
                speak(f"Profile picture of {dp} is downloaded.")
            except Exception as e:
                print(e)
                speak(f"Profile {dp} does not exist.")

        elif 'details of a contact number' in query:
            
            print("Please enter the phone number with country code")
            speak("Please enter the phone number with country code")

            cont=input("Cont.: ")
            cont=phonenumbers.parse(cont)

            tzone = timezone.time_zones_for_number(cont)
            carry = carrier.name_for_number(cont,"en")
            reg = geocoder.description_for_number(cont,"en")
            val = phonenumbers.is_valid_number(cont)
            possib = phonenumbers.is_possible_number(cont)

            print(cont)
            print(f"Time Zone: {tzone}")
            print(f"Carrier: {carry}")
            print(f"Region: {reg}")
            print("Valid Mobile Number: ",val)
            print("Checking possibility of Number: ",possib)

            speak(f"This contact number is from {reg} with time zone is {tzone} and Network carrier is {carry}")

        

        elif 'upload on instagram' in query:
            try:
                uname = "ig_uname"                                      #enter your instagram username
                pwd = "ig_pwd"                                          #enter your instagram password
                speak("Getting image to post.")
                post = input("Enter address of image: ")                #enter image address to be posted
                speak("What is the caption?")
                cap = takeCommand()
                
                bot = Bot()
                bot.login(username=uname , password=pwd)
                bot.upload_photo(post,captioin=cap)
                speak("Feed posted successfully.")

            except Exception as e:
                print(e)
                print("Please try again later.")
                speak("Please try again later.")


        elif 'internet speed' in query:
            test = speedtest.Speedtest()
            down = test.download()
            upload = test.upload()
            print("Download speed : {down}")
            print("Upload speed : {upload}")
            speak(f"Download Speed is {down} and upload speed is {down}")

        elif 'text to handwriting' in query:
            speak("What is the text.")
            text = takeCommand()
            pywhatkit.text_to_handwriting(text,rgb=(0,0,255))

        elif 'website images' in query:
            try:
                def folder_create(images):
                    try:
                        speak("Enter folder name below")
                        folder_name=input("Folder name: ")
                        os.mkdir(folder_name)
                    except:
                        print("Folder exists with the same name")
                        speak("Folder exists with the same name")
                        folder_create()
                    download_images(images,folder_name)

                def download_images(images,folder_name):
                    count=0
                    print(f"Total {len(images)} Images Found!")
                    if len(images) != 0:
                        for i, image in enumerate(images):
                            try:
                                image_link = image["data-srcset"]
                            except:
                                try:
                                    image_link = image["data-src"]
                                except:
                                    try:
                                        image_link = image["data-fallback-src"]
                                    except:
                                        try:
                                            image_link = image["src"]
                                        except:
                                            pass

                            try:
                                r = requests.get(image_link).content
                                try:
                                    r = str(r,'utf=8')

                                except UnicodeDecodeError:
                                    with open(f"{folder_name}/images{i+1}.jpg","wb+") as f:
                                        f.write(r)
                                    count += 1
                            except:
                                pass
                        if count == len(images):
                            print("All images downloaded!")
                        else:
                            print(f"Total {count} images downloaded out of {len(images)}")

                def main(url):
                    r = requests.get(url)
                    soup =BeautifulSoup(r.text,'html.parser')
                    images = soup.findAll('img')
                    folder_create(images)

                speak("Enter website URL to download images.")
                url = input("URL: ")
                main(url)
            
            except Exception as e:
                print(e)
                speak("Please enter valid website URL.")
            






        elif 'lock system' in query:
            speak("Lock")
            ctypes.windll.user32.LockWorkStation()
        
        elif 'logout system' in query:
            os.system("shutdown - l")

        elif 'restart system' in query:
            os.system("shutdown /r /t 1")

        elif 'shutdown system' in query:
            os.system("shutdown /s /t 1")

        elif 'sleep' in query or 'hibernate' in query:
            speak("Hibernating")
            subprocess.call("shutdown / h")

        elif 'log out' in query or 'sign out' in query:
            speak("Make sure all the application are closed before sign-out.")
            time.sleep(5)
            subprocess.call(["shutdown", "/l"])

        elif 'restart window' in query:
            subprocess.call(["shutdown", "/r"])

        elif 'shutdown window' in query:
            subprocess.call('shutdown / p /f')