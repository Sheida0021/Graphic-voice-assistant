import tkinter as tk
import speech_recognition as sr
import pyttsx3
import threading
import datetime
import wikipedia
import os 
import webbrowser
import pyjokes
import pywhatkit as kit
import time 
import subprocess
import wolframalpha
import json
import winshell
import feedparser
import smtplib
import ctypes
import requests
import fileinput
import getpass
import wmi
import datetime
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path
from clint.textui import progress
from selenium import webdriver
from ecapture import ecapture as ec
from bs4 import BeautifulSoup
import win32com.client as wincl
from urllib.request import urlopen
from plyer import notification
from PIL import Image, ImageTk
from threading import Thread
from tkinter import ttk
from tkinter import LEFT, BOTH, SUNKEN




# Constants for custom styling
BG_COLOR = "#D2C6E2"
BUTTON_COLOR = "#F9F4F2"
BUTTON_FONT = ("Arial", 14, "bold")
BUTTON_FOREGROUND = "black"
HEADING_FONT = ("white", 24, "bold")
INSTRUCTION_FONT = ("Helvetica", 14)



r = sr.Recognizer()
engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[1].id)


# Create a Tkinter window for receiving user input
root = tk.Tk()
root.title("Voice Assistant")
root.geometry("500x700")
root.configure(bg=BG_COLOR)

background_image = Image.open(r"C:\\Users\\sheid\\OneDrive\\Desktop\\project_python\\wallpaperflare.com_wallpaper.jpg") # Replace "path/to/your/background_image.jpg" with the actual image file path
background_photo = ImageTk.PhotoImage(background_image)
background_label = ttk.Label(root, image=background_photo)
background_label.place(x=0, y=0, relwidth=1, relheight=1)

f1 = ttk.Frame(root)
f1.pack(pady=100)  # Add some padding to the frame to center it vertically

image2 = Image.open("C:\\Users\\sheid\\OneDrive\\Desktop\\project_python\\p.jpg")  # Replace "path_to_image2.jpg" with the actual path to your image
resized_image = image2.resize((100, 100))
p2 = ImageTk.PhotoImage(resized_image)
l2 = ttk.Label(f1, image=p2, relief=SUNKEN)
l2.pack(side="top", fill="both")

image2 = Image.open("C:\\Users\\sheid\\OneDrive\\Desktop\\project_python\\p.jpg")  # Replace "path_to_image2.jpg" with the actual path to your image
resized_image = image2.resize((100, 100))
p2 = ImageTk.PhotoImage(resized_image)
l2 = ttk.Label(f1, image=p2, relief=SUNKEN)
l2.pack(side="top", fill="both")

# Heading
heading_label = ttk.Label(root, text="Voice Assistant", font=HEADING_FONT, background=BG_COLOR)
heading_label.pack(pady=20)

# Create a label and entry field for user input
f1 = ttk.Frame(root)
f1.pack()
l1 = ttk.Label(f1, text="Enter Your Name", font=INSTRUCTION_FONT, background=BG_COLOR)
l1.pack(side=LEFT, fill=BOTH)
entry = ttk.Entry(f1, width=30)
entry.pack(pady=10)

# Instruction
instruction_label = ttk.Label(root, text="Click the button below to start the Voice Assistant.",font=INSTRUCTION_FONT, background=BG_COLOR)
instruction_label.pack(pady=10)

# Create a button to submit user input
def submit_name():
    name = entry.get()
    root.destroy()
    conversation_window(name)
    

button = ttk.Button(root, text="Start Voice Assistant", command=submit_name,style="VoiceAssistant.TButton")
button.pack(pady=20)

 # Style the button
style = ttk.Style(root)
style.configure("VoiceAssistant.TButton", font=BUTTON_FONT, background=BUTTON_COLOR, foreground=BUTTON_FOREGROUND)

conversation_root = None
# Create a Tkinter window for displaying conversation history
def conversation_window(name):
    global conversation_root
    conversation_root = tk.Tk()
    conversation_root.title("Conversation History")
    conversation_root.geometry("500x500")
    wishMe(name)
    

    # Create a text box for displaying conversation history
    conversation_text = tk.Text(conversation_root, width=70, height=26 )
    conversation_text.pack()

    # Create a label for displaying the user's name
    user_label = tk.Label(conversation_root, text=f"User: {name}")
    user_label.pack()

    # Create a button for speaking to the assistant
    def speak_to_assistant():
        # Start a new thread for speech recognition
        threading.Thread(target=speech_recognition, args=(conversation_text,)).start()
        
    image3 = Image.open(r"C:\\Users\\sheid\\OneDrive\Desktop\\project_python\\123456.png")
    resized_image = image3.resize((50,50 ))
    image3 = ImageTk.PhotoImage(resized_image)

    speak_button = tk.Button(conversation_root, image=image3, command=speak_to_assistant)
    speak_button.pack()

    # Start the conversation window
    conversation_root.mainloop()
            
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

def set_alarm(text):
    try:
        time_str = text.replace("set alarm for ", "")
        alarm_time = datetime.datetime.strptime(time_str, "%I:%M %p")
        current_time = datetime.datetime.now()
        if current_time > alarm_time:
            alarm_time = alarm_time.replace(current_time.year, current_time.month, current_time.day + 1)
        while True:
            time.sleep(1)
            current_time = datetime.datetime.now()
            if current_time.hour == alarm_time.hour and current_time.minute == alarm_time.minute:
                notification.notify(
                    title="Alarm Notification",
                    message="Alarm set for "+time_str+" has been triggered.",
                    app_name="Voice Assistant",
                    timeout=10
                )
                break
    except ValueError:
        notification.notify(
            title="Invalid Time Format",
            message="Please specify the time in 'HH:MM AM/PM' format.",
            app_name="Voice Assistant",
            timeout=10
        )

def speak(audio):
    engine.say(audio)
    engine.runAndWait()

def countdown(n) :
    while n > 0:
        print (n)
        n = n - 1
    if n ==0:
        print('BLAST OFF!')

def wishMe(name):
    hour = int(datetime.datetime.now().hour)
    if hour>=0 and hour<12:
        speak(f"Good Morning {name}!")

    elif hour>=12 and hour<18:
        speak(f"Good Afternoon {name}!")   

    else:
        speak(f"Good Evening {name}!")  

    assname=("Tivo")
    speak("I am your Assistant")
    speak(assname)
    speak("how can i help you?")


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

def speech_recognition(conversation_text):
    recognizer = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening...")
        audio = r.listen(source)
        try:
            text = r.recognize_google(audio, language="en-US")
            print("You said:", text)
            conversation_text.insert(tk.END, f"User: {text}\n")
            response = process_text(text, conversation_text)
            #conversation_text.insert(tk.END, f"Assistant: {response}\n")
            #engine.say(response)
            engine.runAndWait()
        except sr.UnknownValueError:
            print("Sorry, I didn't understand that.")
            

def start(text, conversation_text):
    hour = int(datetime.datetime.now().hour)
    if hour>=0 and hour<12:
        speak("Good Morning !")

    elif hour>=12 and hour<18:
        speak("Good Afternoon !")   

    else:
        speak("Good Evening !")  
    speak("how can i help you?")

def get_weather(city):
    api_key = "f5f13c3f6d997b396795738b674115cc"  # Replace "YOUR_API_KEY" with your OpenWeatherMap API key
    base_url = "http://api.openweathermap.org/data/2.5/weather?"
    complete_url = base_url + "q=" + city + "&appid=" + api_key

    response = requests.get(complete_url)
    data = response.json()

    if data["cod"] != "404":
        main = data["main"]
        temperature = main["temp"] - 273.15
        humidity = main["humidity"]
        weather = data["weather"][0]["description"]

        return f"The temperature in {city} is {temperature:.2f}°C, humidity is {humidity}%, and weather is {weather}."
    else:
        return "City not found."

def stop_voice_assistant():
    global stop_flag
    speak("Stopping the Voice Assistant.")
    stop_flag = True
    conversation_root.destroy()

def process_text(text, conversation_text):
    text = text.lower()
    if 'wikipedia' in text:
            speak('Searching Wikipedia...')
            text = text.replace("wikipedia", "")
            try:
                results = wikipedia.summary(text, sentences=2)
                speak("According to Wikipedia")
                print(results)
                conversation_text.insert(tk.END,results)
                speak(results)
            except wikipedia.exceptions.DisambiguationError as e:
                # Handle disambiguation error (when the search term has multiple possible meanings)
                conversation_text.insert(tk.END,f"There are multiple meanings for '{text}'. Please be more specific.")
                print(f"There are multiple meanings for '{text}'. Please be more specific.")
                conversation_text.insert(tk.END,f"There are multiple meanings for '{text}'. Please be more specific.")
                speak(f"There are multiple meanings for '{text}'. Please be more specific.")
            except wikipedia.exceptions.PageError as e:
                # Handle page not found error (when the search term does not match any Wikipedia page)
                print(f"'{text}' does not match any Wikipedia page. Please try again.")
                conversation_text.insert(tk.END,f"'{text}' does not match any Wikipedia page. Please try again.")
                speak(f"'{text}' does not match any Wikipedia page. Please try again.")
        
    elif "where is" in text:
            text = text.replace("where is", "")
            location = text
            speak("User asked to Locate")
            speak(location)
            conversation_text.insert(tk.END,location)
            webbrowser.open("https://www.google.nl/maps/place/" + location.replace(" ", "+"))

    elif 'open youtube' in text:
            speak("Taking You To Youtube\n")
            conversation_text.insert(tk.END,"Taking You To Youtube\n")
            webbrowser.open("youtube.com")

    elif "open google" in text.lower():
            conversation_text.insert(tk.END,"Taking you to Google\n")
            webbrowser.open("google.com")
            return("Taking you to Google\n")

    elif "change brightness to " in text:
            text=text.replace("change brightness to","")
            conversation_text.insert(tk.END,"change brightness")
            brightness = text 
            c = wmi.WMI(namespace='wmi')
            methods = c.WmiMonitorBrightnessMethods()[0]
            methods.WmiSetBrightness(brightness, 0)

    elif 'open stackoverflow' in text:
            speak("Here you go to Stack Over flow.Happy coding")
            conversation_text.insert(tk.END,"Here you go to Stack Over flow.Happy coding")
            webbrowser.open("stackoverflow.com")   

    elif "github " in text:
            speak("Stackoverflow khola ja rha h")
            conversation_text.insert(tk.END,"Stackoverflow khola ja rha h")
            webbrowser.open("http://github.com")   

    elif 'search' in text : 
            text = text.replace("search", "") 
            speak("searching:")
            conversation_text.insert(tk.END,"searching:")
            webbrowser.open(text)     

    elif 'play music' in text or "play song" in text or "gaana"in text or "song" in text:
            music_dir = "D:\\music"
            username = getpass.getuser()
            music_dir = "D:\\music\\good"
            songs = os.listdir(music_dir)
            print(songs)    
            random=os.startfile(os.path.join(music_dir, songs[1]))

    elif 'the time' in text:
            strTime = datetime.datetime.now().strftime("%H:%M:%S")    
            speak(f", the time is {strTime}")
            conversation_text.insert(tk.END,f", the time is {strTime}")

    elif "samay" in text:
            strTime = datetime.datetime.now().strftime("%H:%M:%S")    
            speak(f"samaye hai {strTime}")
            conversation_text.insert(tk.END,f"samaye hai {strTime}")

    elif 'open VLC' in text:
            codePath = r"C:\\Users\\sheid\\AppData\\Roaming\\Microsoft\Windows\\Start Menu\\Programs\\vlc.exe"
            os.startfile(codePath)
            speak("open VLC")
            conversation_text.insert(tk.END,"open VLC")
        
    elif 'how are you' in text:
                speak("I am fine , Thank you")
                conversation_text.insert(tk.END,"I am fine , Thank you")
                speak("How are you, ")
                conversation_text.insert(tk.END,"How are you, ")

    elif "change name" in text:
            speak("What would you like to call me , ")
            conversation_text.insert(tk.END,"What would you like to call me , ")
            assname = speech_recognition()
            speak("Thanks for naming me")
            conversation_text.insert(tk.END,"Thanks for naming me")

    elif "what's your name" in text or "What is your name" in text:
            speak("My friends call me")
            speak(assname)
            conversation_text.insert(tk.END,"My friends call me",assname)
            print("My friends call me",assname)
            

    elif 'exit' in text:
            speak("Thanks for giving me your time")
            conversation_text.insert(tk.END,"Thanks for giving me your time")
            stop_voice_assistant()
            exit()
            

    elif "who made you" in text or "who created you" in text:
            speak("I have been created by sheida and fatemeh.")
            conversation_text.insert(tk.END,"I have been created by sheida and fatemeh.")

    elif 'joke' in text:
            speak(pyjokes.get_joke())
            conversation_text.insert(tk.END,pyjokes.get_joke())

    elif 'add' in text :
            numbers = [int(word) for word in text.split() if word.isdigit()]
            result = sum(numbers)
            speak(result)
            conversation_text.insert(tk.END, result)
    elif "subtraction" in text:
          numbers = [int(word) for word in text.split() if word.isdigit()]
          result = numbers[0] - numbers[1]
          speak(result)
          conversation_text.insert(tk.END, result)
    elif "multiply" in text:
          numbers = [int(word) for word in text.split() if word.isdigit()]
          result = numbers[0] * numbers[1]
          speak(result)
    elif "divide" in text:
          numbers = [int(word) for word in text.split() if word.isdigit()]
          if numbers[1] != 0:
                result = numbers[0] / numbers[1]
                speak(result)
                conversation_text.insert(tk.END, result)
          else:
                speak("Cannot divide by zero")
                conversation_text.insert(tk.END, "Cannot divide by zero")
           
    elif "who i am" in text:
            speak("If you talk then definately your human.")
            conversation_text.insert(tk.END,"If you talk then definately your human.")

    elif "why you came to world" in text or "reason for you" in text:
            speak("I was created to make your work easier and faster.")
            conversation_text.insert(tk.END,"I was created to make your work easier and faster.")


    elif "who are you" in text:
            speak("I am your virtual assistant created by sheida and fatemeh")
            conversation_text.insert(tk.END,"I am your virtual assistant created by sheida and fatemeh")

    elif 'change background' in text:
            ctypes.windll.user32.SystemParametersInfoW(20, 0, r"C:\\Users\\sheid\\OneDrive\\Desktop\\project_python\\1.jpg" , 0)
            speak("Background changed succesfully")
            conversation_text.insert(tk.END,"Background changed succesfully")

    elif 'open vscode' in text:
            appli= r"C:\\Users\\sheid\\AppData\\Local\\Programs\\Microsoft VS Code\\Visual Studio Code.exe"
            os.startfile(appli)

    elif 'open Google Meet' in text or 'Chrome' in text:
            appli= r"C:\\Users\\sheid\\OneDrive\\Desktop\\Google Meet.exe"
            os.startfile(appli)

    elif 'open Google Chrome' in text or 'Chrome' in text:
            appli= r"C:\\Program Files\\Google\\Chrome\\Application\\Google Chrome.exe"
            os.startfile(appli)

    elif 'Google news' in text:
            try:
                jsonObj = urlopen('''https://newsapi.org/v2/everything?q=apple&from=2024-05-06&to=2024-05-06&sortBy=popularity&apiKey=ba5540ba1ef1446993acbbb8fd52cfb3''')
                data = json.load(jsonObj)
                i = 1
                speak('')
                print('''===============Google News============'''+ '\n')
                conversation_text.insert(tk.END,'''===============Google News============'''+ '\n')
                for item in data['articles'][:5]:  # Limit the number of news items to 5
                    print(str(i) + '. ' + item['title'] + '\n')
                    conversation_text.insert(tk.END,str(i) + '. ' + item['title'] + '\n')
                    print(item['description'] + '\n')
                    speak(str(i) + '. ' + item['title'] + '\n')
                    conversation_text.insert(tk.END,str(i) + '. ' + item['title'] + '\n')
                    i += 1
            except Exception as e:
                    print(str(e))

    elif 'lock window' in text or "system ko lock Karen" in text:
            speak("locking the device")
            conversation_text.insert(tk.END,"locking the device")
            ctypes.windll.user32.LockWorkStation()

    elif 'shutdown system' in text:
            speak("Hold On a Sec! Your system is on its way to shut down")
            conversation_text.insert(tk.END,"Hold On a Sec! Your system is on its way to shut down")
            subprocess.call('shutdown /p /f')
            
    elif 'empty recycle bin' in text:
            winshell.recycle_bin().empty(confirm=False, show_progress=False, sound=True)
            speak("Recycle Bin Recycled")
            conversation_text.insert(tk.END,"Recycle Bin Recycled")

    elif "where is" in text:
            text=text.replace("where is","")
            location = text
            speak("Locating ")
            conversation_text.insert(tk.END,"Locating ")
            speak(location)
            conversation_text.insert(tk.END,location)
            webbrowser.open("https://www.google.nl/maps/place/" + location + "")

    elif "camera" in text or "take a photo" in text:
            ec.capture(0,"Tivo Camera ","img.jpg")

    elif "restart" in text:
            subprocess.call(["shutdown", "/r"])

    elif "hibernate" in text or "sleep" in text:
            speak("Hibernating")
            subprocess.call("shutdown /i /h")

    elif "log off" in text or "sign out" in text:
            speak("Make sure all the application are closed before sign-out")
            conversation_text.insert(tk.END,"Make sure all the application are closed before sign-out")
            time.sleep(5)
            subprocess.call(["shutdown", "/l"])

    elif "countdown of" in text:
            text = int(text.replace("countdown of ",""))
            countdown(text)

    elif "set alarm for" in text:
          set_alarm(text)

    elif 'weather in' in text:
        city = text.split("in", 1)[1].strip()
        weather_info = get_weather(city)
        speak(weather_info)
        conversation_text.insert(tk.END, weather_info) 
    
    elif "wikipedia" in text:
            webbrowser.open("wikipedia.com")

    elif "what is" in text or "who is" in text:
            client= wolframalpha.Client("WOlframe Alpha API key")
            res = client.text(text)
            try:
                print(next(res.results).text)
                speak(next(res.results).text)
                conversation_text.insert(tk.END,next(res.results).text)
            except StopIteration:
                print ("No results")
                conversation_text.insert(tk.END,"No results")

    elif "open Gmail" in text:
            webbrowser.open("https://mail.google.com")
            speak("open Gmail")
            conversation_text.insert(tk.END,"open Gmail")

    elif "open yahoo mail" in text:
            webbrowser.open("https://in.mail.yahoo.com")
            speak("open yahoo mail")
            conversation_text.insert(tk.END,"open yahoo mail")

    elif 'exit' in text:
            speak("thanks for giving your time")
            conversation_text.insert(tk.END,"thanks for giving your time")
            stop_voice_assistant()

# Start the program
root.mainloop()