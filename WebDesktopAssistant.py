import pyttsx3
import datetime
import speech_recognition as sr
import os
import webbrowser
import sys
import time
from random import choice

# Set the console to use UTF-8 encoding
sys.stdout.reconfigure(encoding='utf-8')

# Initialize the text-to-speech engine
engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[0].id)

def speak(audio):
    engine.say(audio)
    engine.runAndWait()

def wishMe():
    hour = int(datetime.datetime.now().hour)
    if 0 <= hour < 12:
        speak("Good morning!")
    elif 12 <= hour < 18:
        speak("Good afternoon!")
    else:
        speak("Good evening!")
    speak("I am , your desktop assistant. How may I assist you today?")

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
        print("Say that again, please...")
        return "None"
    return query

# Function to open standard applications
def openApp(app_name):
    app_paths = {
        'excel': "C:\\Program Files\\Microsoft Office\\root\\Office16\\EXCEL.EXE",
        'powerpoint': "C:\\Program Files\\Microsoft Office\\root\\Office16\\POWERPNT.EXE",
        'word': "C:\\Program Files\\Microsoft Office\\root\\Office16\\WINWORD.EXE",
        'notepad': "C:\\Windows\\System32\\notepad.exe",
        'calculator': "C:\\Windows\\System32\\calc.exe",
        'paint': "C:\\Windows\\System32\\mspaint.exe",
        'browser': "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe",
        'edge': "C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe"
    }

    app_path = app_paths.get(app_name.lower())
    if app_path and os.path.exists(app_path):
        os.startfile(app_path)
        speak(f"Opening {app_name}")
        speak(f"{app_name} is now open.")
    else:
        speak(f"{app_name} is not present on this computer.")

# Function to search for and open images on the desktop
def searchAndOpenImages(query):
    desktop_path = os.path.expanduser("~/Desktop")
    image_extensions = ['.png', '.jpg', '.jpeg', '.bmp', '.gif', '.webp']
    image_files = [f for f in os.listdir(desktop_path) if os.path.splitext(f)[1].lower() in image_extensions]

    if image_files:
        speak("Here are the images I found on your desktop.")
        for image in image_files:
            image_path = os.path.join(desktop_path, image)
            os.startfile(image_path)  # Opens the image with the default viewer
            time.sleep(2)  # Waits 2 seconds between images
    else:
        speak("No images found on the desktop.")

if __name__ == "__main__":
    wishMe()

    while True:
        query = takeCommand().lower()

        if 'search' in query or 'google' in query:
            speak("Searching Google...")
            query = query.replace("search", "").replace("google", "")
            webbrowser.open(f"https://www.google.com/search?q={query}")
            speak(f"Here are the results for {query} on Google.")

        elif 'open youtube' in query:
            webbrowser.open("https://www.youtube.com")
            speak("Opening YouTube")
        elif 'open google' in query:
            webbrowser.open("https://www.google.com")
            speak("Opening Google")
        elif 'the time' in query:
            strTime = datetime.datetime.now().strftime("%H:%M:%S")
            speak(f"The time is {strTime}")
        elif 'open code' in query:
            codepath = "C:\\Users\\Hp\\AppData\\Local\\Programs\\Microsoft VS Code\\Code.exe"
            os.startfile(codepath)
            speak("Opening Visual Studio Code")
        elif 'open excel' in query:
            openApp('excel')
        elif 'open powerpoint' in query:
            openApp('powerpoint')
        elif 'open word' in query:
            openApp('word')
        elif 'open notepad' in query:
            openApp('notepad')
        elif 'open calculator' in query:
            openApp('calculator')
        elif 'open paint' in query:
            openApp('paint')
        elif 'open browser' in query:
            openApp('browser')
        elif 'open edge' in query:
            openApp('edge')
        elif 'show images' in query:
            searchAndOpenImages(query)
        elif 'show slideshow' in query:
            searchAndOpenImages('slideshow')
        elif 'exit' in query or 'quit' in query:
            speak("Goodbye!")
            break

        else:
            speak("I didn't catch that. Can you repeat?")
