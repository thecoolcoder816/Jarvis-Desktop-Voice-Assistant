import pyttsx3
import datetime
import speech_recognition as sr
import wikipedia
import webbrowser as wb
import os
import random
import pyautogui
import pyjokes
import math
import re
import platform

# For unit conversion
from typing import Tuple

# For volume control (Windows only, Linux, macOS)
if platform.system() == "Windows":
    try:
        from ctypes import cast, POINTER
        from comtypes import CLSCTX_ALL
        from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
    except ImportError:
        pass  # pycaw not installed

engine = pyttsx3.init()
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[1].id)  
engine.setProperty('rate', 150)
engine.setProperty('volume', 1)

def speak(audio) -> None:
    engine.say(audio)
    engine.runAndWait()

def time() -> None:
    current_time = datetime.datetime.now().strftime("%I:%M:%S %p")
    speak("The current time is")
    speak(current_time)
    print("The current time is", current_time)

def date() -> None:
    now = datetime.datetime.now()
    speak("The current date is")
    speak(f"{now.day} {now.strftime('%B')} {now.year}")
    print(f"The current date is {now.day}/{now.month}/{now.year}")

def wishme() -> None:
    speak("Welcome back, sir!")
    print("Welcome back, sir!")

    hour = datetime.datetime.now().hour
    if 4 <= hour < 12:
        speak("Good morning!")
        print("Good morning!")
    elif 12 <= hour < 16:
        speak("Good afternoon!")
        print("Good afternoon!")
    elif 16 <= hour < 24:
        speak("Good evening!")
        print("Good evening!")
    else:
        speak("Good night, see you tomorrow.")

    assistant_name = load_name()
    speak(f"{assistant_name} at your service. Please tell me how may I assist you.")
    print(f"{assistant_name} at your service. Please tell me how may I assist you.")

def screenshot() -> None:
    img = pyautogui.screenshot()
    img_path = os.path.expanduser("~\\Pictures\\screenshot.png")
    img.save(img_path)
    speak(f"Screenshot saved as {img_path}.")
    print(f"Screenshot saved as {img_path}.")

def takecommand() -> str:
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening...")
        r.pause_threshold = 1

        try:
            audio = r.listen(source, timeout=5)
        except sr.WaitTimeoutError:
            speak("Timeout occurred. Please try again.")
            return None

    try:
        print("Recognizing...")
        query = r.recognize_google(audio, language="en-in")
        print(query)
        return query.lower()
    except sr.UnknownValueError:
        speak("Sorry, I did not understand that.")
        return None
    except sr.RequestError:
        speak("Speech recognition service is unavailable.")
        return None
    except Exception as e:
        speak(f"An error occurred: {e}")
        print(f"Error: {e}")
        return None

def play_music(song_name=None) -> None:
    song_dir = os.path.expanduser("~\\Music")
    songs = os.listdir(song_dir)

    if song_name:
        songs = [song for song in songs if song_name.lower() in song.lower()]

    if songs:
        song = random.choice(songs)
        os.startfile(os.path.join(song_dir, song))
        speak(f"Playing {song}.")
        print(f"Playing {song}.")
    else:
        speak("No song found.")
        print("No song found.")

def set_name() -> None:
    speak("What would you like to name me?")
    name = takecommand()
    if name:
        with open("assistant_name.txt", "w") as file:
            file.write(name)
        speak(f"Alright, I will be called {name} from now on.")
    else:
        speak("Sorry, I couldn't catch that.")

def load_name() -> str:
    return "Jarvis"

def search_wikipedia(query):
    try:
        speak("Searching Wikipedia...")
        result = wikipedia.summary(query, sentences=3)
        speak(result)
        print(result)
    except wikipedia.exceptions.DisambiguationError:
        speak("Multiple results found. Please be more specific.")
    except Exception:
        speak("I couldn't find anything on Wikipedia.")

# ----------- Added Feature: Calculator -----------

def calculate_expression(expression: str) -> None:
    try:
        allowed = "0123456789+-*/(). "
        safe_expr = "".join([ch for ch in expression if ch in allowed])
        if safe_expr.strip() == "":
            speak("I couldn't understand the calculation.")
            return
        result = eval(safe_expr, {"__builtins__": None}, {})
        speak(f"The result is {result}")
        print(f"Calculation: {safe_expr} = {result}")
    except Exception:
        speak("Sorry, I couldn't calculate that.")
        print("Error in calculation.")

# ----------- Cross-platform Volume Control -----------

def set_volume(level: int) -> None:
    level = max(0, min(level, 100))
    sys = platform.system()
    if sys == "Windows":
        try:
            devices = AudioUtilities.GetSpeakers()
            interface = devices.Activate(
                IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
            volume = cast(interface, POINTER(IAudioEndpointVolume))
            volume.SetMasterVolumeLevelScalar(level / 100, None)
            speak(f"Volume set to {level} percent.")
            print(f"Volume set to {level}%")
        except Exception:
            speak("Failed to set volume.")
            print("Volume control failed.")
    elif sys == "Linux":
        try:
            import subprocess
            subprocess.run(["amixer", "sset", "Master", f"{level}%"], check=True)
            speak(f"Volume set to {level} percent.")
            print(f"Volume set to {level}%")
        except Exception:
            speak("Failed to set volume on Linux.")
            print("Volume control failed on Linux.")
    elif sys == "Darwin":  # macOS
        try:
            os.system(f"osascript -e 'set volume output volume {level}'")
            speak(f"Volume set to {level} percent.")
            print(f"Volume set to {level}%")
        except Exception:
            speak("Failed to set volume on macOS.")
            print("Volume control failed on macOS.")
    else:
        speak("Sorry, volume control is not supported on your operating system.")
        print("Volume control not supported on this OS.")

# ----------- Added Feature: Unit Converter -----------

def convert_units(query: str) -> None:
    conversions = {
        ("kilometers", "miles"): lambda x: x * 0.621371,
        ("miles", "kilometers"): lambda x: x / 0.621371,
        ("meters", "feet"): lambda x: x * 3.28084,
        ("feet", "meters"): lambda x: x / 3.28084,
        ("celsius", "fahrenheit"): lambda x: (x * 9/5) + 32,
        ("fahrenheit", "celsius"): lambda x: (x - 32) * 5/9,
        ("kilograms", "pounds"): lambda x: x * 2.20462,
        ("pounds", "kilograms"): lambda x: x / 2.20462,
        ("grams", "ounces"): lambda x: x * 0.035274,
        ("ounces", "grams"): lambda x: x / 0.035274,
        ("liters", "gallons"): lambda x: x * 0.264172,
        ("gallons", "liters"): lambda x: x / 0.264172,
    }
    pattern = r"convert (\d+(?:\.\d+)?) (\w+) to (\w+)"
    match = re.search(pattern, query)
    if match:
        amount = float(match.group(1))
        from_unit = match.group(2).lower()
        to_unit = match.group(3).lower()
        func = conversions.get((from_unit, to_unit))
        if func:
            result = func(amount)
            speak(f"{amount} {from_unit} is {result:.2f} {to_unit}")
            print(f"{amount} {from_unit} = {result:.2f} {to_unit}")
        else:
            speak("Sorry, I can't convert those units.")
    else:
        speak("Please say, for example: convert 10 kilometers to miles.")

if __name__ == "__main__":
    wishme()

    while True:
        query = takecommand()
        if not query:
            continue

        if "time" in query:
            time()
            
        elif "date" in query:
            date()

        elif "wikipedia" in query:
            query = query.replace("wikipedia", "").strip()
            search_wikipedia(query)

        elif "play music" in query:
            song_name = query.replace("play music", "").strip()
            play_music(song_name)

        elif "open youtube" in query:
            wb.open("youtube.com")
            
        elif "open google" in query:
            wb.open("google.com")

        elif "change your name" in query:
            set_name()

        elif "screenshot" in query:
            screenshot()
            speak("I've taken screenshot, please check it")

        elif "tell me a joke" in query:
            joke = pyjokes.get_joke()
            speak(joke)
            print(joke)

        elif "shutdown" in query:
            speak("Shutting down the system, goodbye!")
            os.system("shutdown /s /f /t 1")
            break
            
        elif "restart" in query:
            speak("Restarting the system, please wait!")
            os.system("shutdown /r /f /t 1")
            break
            
        elif "offline" in query or "exit" in query:
            speak("Going offline. Have a good day!")
            break

        # Feature: Calculator
        elif "calculate" in query:
            expr = query.replace("calculate", "").replace("what is", "").strip()
            calculate_expression(expr)

        # Feature: Volume Control
        elif "set volume" in query:
            match = re.search(r"set volume (\d+)", query)
            if match:
                level = int(match.group(1))
                set_volume(level)
            else:
                speak("Please specify a volume between 0 and 100.")

        # Feature: Unit Converter
        elif "convert" in query and "to" in query:
            convert_units(query)