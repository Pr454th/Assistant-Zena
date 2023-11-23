import subprocess
import wolframalpha
import wikiquote
import pyttsx3
import json
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
import fileinput
import getpass
import pyautogui
import wmi
import os
import random
from pathlib import Path
from clint.textui import progress
from selenium import webdriver
from ecapture import ecapture as ec
from bs4 import BeautifulSoup
import win32com.client as wincl
from urllib.request import urlopen
from word2number import w2n
from PIL import Image
from dotenv import load_dotenv

load_dotenv()

engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[1].id)

DIRECTORIES = {
    "HTML": [".html5", ".html", ".htm", ".xhtml"],
    "IMAGES": [".jpeg", ".jpg", ".tiff", ".gif", ".bmp", ".png", ".bpg", "svg",
               ".heif", ".psd"],
    "VIDEOS": [".avi", ".flv", ".wmv", ".mov", ".mp4", ".webm", ".vob", ".mng",
               ".qt", ".mpg", ".mpeg", ".3gp", ".mkv"],
    "DOCUMENTS": [".oxps", ".epub", ".pages", ".docx", ".doc", ".fdf", ".ods",
                  ".odt", ".pwi", ".xsn", ".xps", ".dotx", ".docm", ".dox",
                  ".rvg", ".rtf", ".rtfd", ".wpd", ".xls", ".xlsx", ".ppt",
                  "pptx"],
    "ARCHIVES": [".a", ".ar", ".cpio", ".iso", ".tar", ".gz", ".rz", ".7z",
                 ".dmg", ".rar", ".xar", ".zip"],
    "AUDIO": [".aac", ".aa", ".aac", ".dvf", ".m4a", ".m4b", ".m4p", ".mp3",
              ".msv", "ogg", "oga", ".raw", ".vox", ".wav", ".wma"],
    "PLAINTEXT": [".txt", ".in", ".out"],
    "PDF": [".pdf"],
    "PYTHON": [".py",".pyi"],
    "XML": [".xml"],
    "EXE": [".exe"],
    "SHELL": [".sh"]
}
FILE_FORMATS = {file_format: directory
                for directory, file_formats in DIRECTORIES.items()
                for file_format in file_formats}


def speak(audio):
    engine.say(audio)
    engine.runAndWait()

def get_number_from_text(text):
    try:
        return w2n.word_to_num(text)
    except ValueError:
        return None

def countdown(n) :
    while n > 0:
        print (n)
        n = n - 1
    if n ==0:
        print('BLAST OFF!')

wapi=os.getenv("WEATHER_API_KEY")
napi=os.getenv("NEWS_API_KEY")
mailid=os.getenv("MAIL_ID")
mailpass=os.getenv("MAIL_PASSCODE")
capi=os.getenv("CALCULATOR_API_KEY")

def wishMe():
    hour = int(datetime.datetime.now().hour)
    if hour>=0 and hour<12:
        speak("Good Morning Sir!")

    elif hour>=12 and hour<18:
        speak("Good Afternoon Sir!")   

    else:
        speak("Good Evening Sir!")  

    assname=("Zena 1 point o")
    speak("I am your Assistant")
    speak(assname)

def usrname():
    speak("What should i call you sir")
    uname=takeCommandname()
    speak("Welcome Mister")
    speak(uname)
    print("#####################")
    print("Welcome Mr.",uname)
    print("#####################")

def quotaton():
    speak(wikiquote.quote_of_the_day())
    print(wikiquote.quote_of_the_day())

def takeCommand():
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
        print("Unable to Recognizing your voice.")  
        return "None"
    return query

def takeCommandname():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        try:
            print("Username...")
            print(wapi)
            print(napi)
            print(mailid)
            print(mailpass)
            print(capi)
            r.pause_threshold = 1
            # r.adjust_for_ambient_noise(source,duration = 1)
            audio = r.listen(source,timeout=5)
        except:
            print("Unable to Recognizing your name.")
            takeCommandname()
            return "None"
    try:
        print("Trying to Recognizing Name...")    
        query = r.recognize_google(audio, language='en-in')
        print(f"User said: {query}\n")

    except Exception as e:
        print(e)    
        print("Unable to Recognizing your name.")
        takeCommandname()  
        return "None"
    return query

def takeCommandmessage():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("Enter Your Message")
        r.pause_threshold = 1
        audio = r.listen(source)

    try:
        query = r.recognize_google(audio, language='en-in')
        print(f'Message to be sent is : {query}\n')

    except Exception as e:
        print (e)
        print("Unable to recognize your message")
        print("Check your Internet Connectivity")
    return query

def takeCommanduser():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("Name of User or Group")
        r.pause_threshold = 1
        audio = r.listen(source)

    try:
        query = r.recognize_google(audio, language='en-in')
        print(f'Client to whome message is to be sent is : {query}\n')

    except Exception as e:
        print (e)
        print("Unable to recognize Client name")
        speak("Unable to recognize Client Name")
        print("Check your Internet Connectivity")
    return query

def takeCommandcontent():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("What Should i say, sir")
        r.pause_threshold = 1
        audio = r.listen(source)

    try:
        query = r.recognize_google(audio, language='en-in')
        print(f'Message to be sent is: {query}\n')

    except Exception as e:
        print (e)
        print("Unable to recognize")
    return query

def organize():
    for entry in os.scandir():
        if entry.is_dir():
            continue
        file_path = Path(entry.name)
        file_format = file_path.suffix.lower()
        if file_format in FILE_FORMATS:
            directory_path = Path(FILE_FORMATS[file_format])
            directory_path.mkdir(exist_ok=True)
            file_path.rename(directory_path.joinpath(file_path))
    try:
        os.mkdir("OTHER")
    except:
        pass
    for dir in os.scandir():
        try:
            if dir.is_dir():
                os.rmdir(dir)
            else:
                os.rename(os.getcwd() + '/' + str(Path(dir)), os.getcwd() + '/OTHER/' + str(Path(dir)))
        except:
            pass

def sendEmail(to, content):
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.ehlo()
    server.starttls()
    server.login(mailid, mailpass)
    server.sendmail(mailid, to, content)
    server.close()


if __name__ == '__main__':
    clear = lambda: os.system('cls')
    clear()
    wishMe()
    usrname()
    # speak("Can i tell you a quote of day")
    # useropt=takeCommand().lower()
    # if 'yes' in useropt or 'sure' in useropt:
    #     quotaton()
    # else:
    #     speak("Taking you to command function")

    speak("How can i Help you, Sir")
    while True:
        query = takeCommand().lower()
        assname=("Zena 1 point o")
        if 'wikipedia' in query and "search" in query:
            speak('Searching Wikipedia...')
            query = query.replace("wikipedia", "").replace("search","")
            results = wikipedia.summary(query, sentences=3)
            speak("Answer From Wikipedia")
            print(results)
            speak(results)

        elif "Good Morning" in query:
            speak("A warm" +query)
            speak("How are you Mister")
            speak(assname)

        elif "wikipedia" in query and "hindi" in query:
            speak('Searching Wikipedia...')
            query = query.replace("wikipedia", "")
            query = query.replace("hindi", "")
            results = wikipedia.summary(query, sentences=3)
            speak("According to Wikipedia")
            r = sr.Recognizer()
            results = r.recognize_google(results, language='hi')
            print(results)
            speak(results)

        elif ('search' in query or 'open' in query) and 'youtube' in query:
            search_query = query.replace('open', '').replace('youtube', '').replace('in','').strip()
            speak(f"Searching for {search_query} on YouTube")
            webbrowser.open(f"https://www.youtube.com/results?search_query={search_query}")
            speak("Taking You To Youtube\n")

        elif 'open google' in query:
            speak("Taking you to Google\n")
            webbrowser.open("google.com")

        elif "change brightness to " in query:
            query=query.replace("change brightness to","")
            brightness = query 
            c = wmi.WMI(namespace='wmi')
            methods = c.WmiMonitorBrightnessMethods()[0]
            methods.WmiSetBrightness(brightness, 0)

        elif "Organize Files" in query:
            organize()

        elif 'open stackoverflow' in query or 'open stack overflow' in query or 'open Stack overflow' in query:
            speak("Here you go to Stack Over flow.Happy coding")
            webbrowser.open("stackoverflow.com")    

        elif "send a whatsaap message" in query or "send a WhatsApp message" in query:
            driver = webdriver.Firefox(executable_path="./geckodriver.exe")
            driver.get('https://web.whatsapp.com/')
            speak("Scan QR code before proceding")
            tim=10
            time.sleep(tim)
            speak("Enter Name of Group or User")
            name = takeCommanduser()
            speak("Enter Your Message")
            msg = takeCommandmessage()
            count = 1
            user = driver.find_element_by_xpath('//span[@title = "{}"]'.format(name))
            user.click()
            msg_box = driver.find_element_by_class_name('_3u328')
            for i in range(count):
                msg_box.send_keys(msg)
                button = driver.find_element_by_class_name('_3M-N-')
                button.click()
                
        # elif "search" in query and ("stackoverflow " in query or "stack overflow" in query or "stackover flow" in query):
        #     search_query = query.replace('search', '').replace('stackoverflow', '').replace('stack overflow', '').replace('stackover flow', '').replace(' ','+')
        #     webbrowser.open(f"stackoverflow.com/search?q={search_query}")

        elif 'play music' in query or "play song" in query:
            #music_dir = "G:\\Song"
            username = getpass.getuser()
            print(username)
            music_dir = "C:\\Users\\"+username+"\\Music"
            songs = os.listdir(music_dir)
            randomSong=songs[random.randint(1,len(songs)-1)]
            print(os.path.join(music_dir, randomSong))
            random=os.startfile(os.path.join(music_dir, randomSong),"open")

        elif 'play movie' in query:
            username = getpass.getuser()
            print(username)
            movie_dir = "C:\\Users\\" + username + "\\Videos"
            movies = os.listdir(movie_dir)
            
            # Print the movie names
            for i, movie in enumerate(movies):
                print(f"{i+1}. {movie}")
                speak(f"{i+1}. {movie}")
            
            # Ask the user to input the movie number
            speak("Please enter the movie number:")
            movie_number = takeCommand()
            movie_number = movie_number.replace("number", "").strip()
            movie_number = int(movie_number)
            # Play the selected movie
            if movie_number >= 1 and movie_number <= len(movies):
                movie_name = movies[movie_number - 1]
                print(os.path.join(movie_dir, movie_name))
                os.startfile(os.path.join(movie_dir, movie_name), "open")
            else:
                print("Invalid movie number")
                speak("Invalid movie number")

        elif 'the time' in query:
            strTime = datetime.datetime.now().strftime("%H:%M:%S")    
            speak(f"Sir, the time is {strTime}")

        elif "open Gmail" in query or "open gmail" in query:
            webbrowser.open("https://gmail.google.com/mail/u/0/#inbox")

        elif "open yahoo mail" in query:
            webbrowser.open("https://in.mail.yahoo.com")

        elif 'open' in query:
            applications = {
                "notepad.exe": ["notepad", "note pad", "Notepad", "Note Pad"],
                "firefox.exe": ["firefox", "mozilla", "browser", "Firefox", "Mozilla", "Browser"],
                "msedge.exe": ["edge", "ms edge", "msedge", "edge app", "Edge", "MS Edge", "Edge App"],
                "winword.exe": ["word", "microsoft word", "ms word", "Word", "Microsoft Word", "MS Word"],
                "excel.exe": ["excel", "microsoft excel", "ms excel", "Excel", "Microsoft Excel", "MS Excel"],
                "powerpnt.exe": ["powerpoint", "microsoft powerpoint", "ms powerpoint", "Powerpoint", "Microsoft Powerpoint", "MS Powerpoint"],
                "outlookcal:": ["calendar", "calender", "schedule", "Calendar", "Calender", "Schedule"],
                "outlook.exe": ["email", "mail", "outlook", "Email", "Mail", "Outlook"],
                "explorer.exe": ["explorer", "file explorer", "Explorer", "File Explorer","file manager","File Manager","file Manager"],
                "ms-settings:": ["settings", "control panel", "Settings", "Control Panel"],
                "ms-photos:": ["photos", "pictures", "Photos", "Pictures"],
                "ms-clock:": ["clock", "alarm", "timer", "Clock", "Alarm", "Timer"],
                "ms-windows-store:": ["store", "windows store", "Store", "Windows Store"],
                "microsoft.teams": ["teams", "microsoft teams", "Teams", "Microsoft Teams"],
                "mspaint.exe": ["paint", "ms paint", "Paint", "MS Paint"],
                "calc.exe": ["calculator", "calc", "Calculator", "Calc"],
                "wt.exe": ["terminal", "cmd", "Terminal", "CMD"],
            }
            application_name = query.replace("open", "").strip()
            for app in applications:
                if application_name in applications[app]:
                    application_name = app
                    break
            try:
                subprocess.run(["start", application_name], shell=True)
                print(f"Opening {application_name}")
            except Exception as e:
                print(f"Error: {e}")

        elif 'take' in query and 'screenshot' in query:
            screen_width, screen_height = pyautogui.size()
            screenshot = pyautogui.screenshot()
            ssName=random.randint(11,999)
            screenshot.save("s" + str(ssName) + ".png")
            speak("Screenshot Taken")
            time.sleep(2)
            speak("Do you want to see the screenshot")
            query=takeCommand()
            if ('yes' in query) or ('sure' in query):
                last_ss = Image.open("s" + str(ssName) + ".png")

        elif 'send'in query and 'email' in query:
            try:
                speak("What should I say?")
                content = takeCommandcontent()
                speak("What should be the subject of this email")
                subject = takeCommand()
                speak("Enter the email address of recipient")
                to = takeCommand()    
                to=to.replace("at","@")
                to=to.replace(" ","")
                to=to.lower()
                sendEmail(to, content)
                speak("Email has been sent!")
            except Exception as e:
                print(e)
                speak("I am not able to send this email")

        elif 'how are you' in query:
        	speak("I am fine , Thank you")
        	speak("How are you, Sir")

        elif "change my name to" in query:
            query=query.replace("change my name to","")
            assname=query

        elif "what's your name" in query or "What is your name" in query:
            speak("My friends call me")
            speak(assname)
            print("My friends call me",assname)

        elif 'exit' in query:
            speak("Thanks for giving me your time")
            exit()

        elif 'joke' in query:
            speak(pyjokes.get_joke())

        elif "calculate" in query:
            app_id = capi
            client = wolframalpha.Client(app_id)
            indx = query.lower().split().index('calculate')
            query = query.split()[indx + 1:]
            res = client.query(' '.join(query))
            answer = next(res.results).text
            print("The answer is " + answer)
            speak("The answer is " + answer)

        elif 'search' in query or 'play' in query: 
            query = query.replace("search", "") 
            query = query.replace("play", "")
            webbrowser.open(query) 

        elif 'change' in query and 'background' in query:
            ctypes.windll.user32.SystemParametersInfoW(20, 0, "C:\\Users\\USER\\Pictures\\Saved Pictures\\wp.jpg" , 0)
            speak("Background changed succesfully")

        elif 'google news' in query:
            try:
                jsonObj = urlopen(f'''https://newsapi.org/v2/top-headlines?sources=google-news-in&apiKey={napi}''')
                data = json.load(jsonObj)
                i = 1
                speak('')
                print('''===============Google News============'''+ '\n')
                for item in data['articles']:
                    print(str(i) + '. ' + item['title'] + '\n')
                    print(item['description'] + '\n')
                    speak(str(i) + '. ' + item['title'] + '\n')
                    i += 1
            except Exception as e:
                    print(str(e))

        elif "bbc news" in query:
            try:
                main_url = f" https://newsapi.org/v1/articles?source=bbc-news&sortBy=top&apiKey={napi}"
                open_bbc_page = requests.get(main_url).json() 
                article = open_bbc_page["articles"] 
                results = [] 
                for ar in article: 
                    results.append(ar["title"]) 
                for i in range(len(results)): 
                    print(i + 1, results[i])
            except Exception as e:
                print(str(e))

        elif 'news' in query:
            try:
                jsonObj = urlopen(f'''https://newsapi.org/v1/articles?source=the-times-of-india&sortBy=top&apiKey={napi}''')
                data = json.load(jsonObj)
                i = 1
                speak('here are some top news from the times of india')
                print('''===============TIMES OF INDIA============'''+ '\n')
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
            speak("Hold On a Sec! Your system is on its way to shut down")
            subprocess.call('shutdown /p /f')
            
        elif 'empty recycle bin' in query:
            winshell.recycle_bin().empty(confirm=False, show_progress=False, sound=True)
            speak("Recycle Bin Recycled")

        elif "don't listen" in query or "stop listening" in query:
            speak("for how much time you want to stop Zena from listening commands")
            a=int(takeCommand())
            time.sleep(a)
            print(a)

        elif "where is" in query:
            query=query.replace("where is","")
            location = query
            speak("Locating ")
            speak(location)
            webbrowser.open("https://www.google.nl/maps/place/" + location + "")

        elif "camera" in query or "take a photo" in query:
            ec.capture(0,"Zena Camera ","img.jpg")

        elif "restart" in query:
            subprocess.call(["shutdown", "/r"])

        elif "hibernate" in query or "sleep" in query:
            speak("Hibernating")
            subprocess.call("shutdown /i /h")

        elif "log off" in query or "sign out" in query:
            speak("Make sure all the application are closed before sign-out")
            time.sleep(5)
            subprocess.call(["shutdown", "/l"])

        elif "countdown of" in query:
            query = int(query.replace("countdown of ",""))
            countdown(query)

        elif "write a note" in query:
            speak("What should i write , sir")
            note= takeCommand()
            file = open('Zena.txt','w')
            speak("Sir, Should i include date and time")
            snfm = takeCommand()
            if 'yes' in snfm or 'sure' in snfm:
                strTime = datetime.datetime.now().strftime("%H:%M:%S")
                file.write(strTime)
                file.write(" :- ")
                file.write(note)
            else:
                file.write(note)
        
        elif "show" in query and "note" in query:
            speak("Showing Notes")
            file = open("Zena.txt", "r") 
            print(file.read())
            speak(file.read(6))

        elif "Zena" in query:
            wishMe()
            speak("Zena 1 point o in your service Mister")
            speak(assname)

        elif "weather" in query:
            api_key = wapi
            base_url = "http://api.openweathermap.org/data/2.5/forecast?"
            print("City name : ")
            city_name=takeCommand()
            complete_url = base_url + "&q=" + city_name+ "&appid=" + api_key
            print(complete_url)
            response = requests.get(complete_url) 
            x = response.json() 
            if x["cod"] != "404": 
                y = x["list"][0]["weather"] 
                z=x["list"][0]["main"]
                current_temperature = z["temp"]
                current_pressure = z["pressure"]
                current_humidiy = z["humidity"]
                current = y[0]["main"] 
                status = y[0]["description"]
                speak(" Temperature (in kelvin unit) = " +str(current_temperature)+"\n atmospheric pressure (in hPa unit) ="+str(current_pressure) +"\n humidity (in percentage) = " +str(current_humidiy) ) 
                speak("Current Weather Status :")
                speak(current)
                speak("Description :")
                speak(status)
            else: 
                speak(" City Not Found ") 

        elif "wikipedia" in query:
            webbrowser.open("wikipedia.com")

        elif "what is" in query or "who is" in query:
            client= wolframalpha.Client("YXEL57-8A48YAHKE8")
            res = client.query(query)
            try:
                print(next(res.results).text)
                speak(next(res.results).text)
            except StopIteration:
                print ("No results")