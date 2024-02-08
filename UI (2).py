import cv2
import pyttsx3 as pyt
import speech_recognition as sr
import datetime as dt
import os
import wikipedia
import webbrowser
import pywhatkit
import requests
from bs4 import BeautifulSoup
from googlesearch import search
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
import math
import random
import time
import schedule
import speedtest
import pyautogui
from AppOpener import open, close
from tkinter import *
from tkinter import messagebox
import tkinter
import glob

r = sr.Recognizer()
mymic = sr.Microphone(device_index=1)

engine = pyt.init()
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[0].id)
engine.setProperty('rate', 190)
engine.setProperty('volume', 0.7)

cap = cv2.VideoCapture("vid.gif")
count = 0


def listen():
    rec = sr.Recognizer()
    with sr.Microphone() as source:
        speak("listening")
        print("Listening....")
        cv2.putText(frame, "Listening...", (100, 80), cv2.FONT_HERSHEY_PLAIN, 1.5, (255, 255, 255), 2)
        audio = rec.listen(source, phrase_time_limit=15)
    try:
        print("Recognizing...")
        query = rec.recognize_google(audio, language='en-in')
        query = query.lower()
        print("You said:", query, "\n")
    except Exception as e:
        speak(" Say that again human")
        print("Say that again human")
        return "None"
    return query


def speak(audio):
    engine.say(audio)
    engine.runAndWait()


def joke():
    myjokes = ["bored of being bored because being bored is boring....ha ha ha ha ha ha",
               "What do you call an Englishman with an IQ of 50?  Colonel, sir.....ha ha ha ha ha ha",
               "They say an Englishman laughs three times at a joke. The first timewhen everybody gets it, the second a week later when he thinks he getsit, the third time a month later when somebody explains it to him.....ha ha ha ha ha ha",
               "Why do cows wear bells?Because their horns don't work....ha ha ha ha ha ha",
               "my life....ha ha ha ha ha ha",
               "your life....ha ha ha ha ha ha",
               "What did the buffalo say to his son when leaving for college ? bison....ha ha ha ha ha ha",
               " I visited my new friend in his flat. He told me to make myself at home. So I threw him out. I hate having visitors....ha ha ha ha ha ha",
               ]
    randomjoke = random.choice(myjokes)
    print(randomjoke)
    speak(randomjoke)


def motivation():
    moti = ["its a shame for a man to grow old without seeing the beauty and strength his body is capable of",
            "one whos ideal is mortal will die when his ideal dies , but when ones ideal is immortal he himself must become immortal to attain it",
            "pain is only in the mind",
            "Success is not how high you have climbed, but how you make a positive difference to the world.",
            "Never lose hope. Storms make people stronger and never last forever.",
            "strong men make great times,great times make weak men,weak men make tough times,tough times make strong men",
            "All our dreams can come true, if we have the courage to pursue them.",
            "The best time to plant a tree was 20 years ago. The second best time is now.",
            "If people are doubting how far you can go, go so far that you can’t hear them anymore.",
            "Everything you can imagine is real."]
    randommoti = random.choice(moti)
    print(randommoti)
    speak(randommoti)


def speed():
    st = speedtest.Speedtest()
    d_speed = st.download()
    u_speed = st.upload()
    KB = 1024
    MB = KB * 1024
    d_mb = int(d_speed / MB)
    u_mb = int(u_speed / MB)
    print(f"Download speed : {d_mb} MBPS")
    speak(f"Your download speed is {d_mb} MB per second")
    print(f"Upload speed : {u_mb} MBPS")
    speak(f"Your upload speed is {u_mb} MB per second")


def about():
    speak("I am Pikolo....a Voice assistant. I am still learning and open to criticism ")


def func():
    speak("i can ")
    speak(
        "open applications , open websites , search things in google, play youtube videos , tell time and day , tell weather, tell jokes ,fun facts , motivational quotes , control system volume")
    speak("what do you want to do ?")


def hydrate():
    print("Drink water human...stay hydrated ")
    speak("Drink water human...stay hydrated ")


def takebk():
    print("take a break human...you deserve it")
    speak("take a break human...you deserve it")


def remind():
    speak("what you want me to remind ?")
    thing = listen()
    print(thing)


def increse_vol():
    devices = AudioUtilities.GetSpeakers()
    interface = devices.Activate(
        IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
    volume = cast(interface, POINTER(IAudioEndpointVolume))

    # Get current volume
    currentVolumeDb = volume.GetMasterVolumeLevel()
    if currentVolumeDb > -5.409796714782715:
        print("volume is more than 70 percent...please keep it below 70")
        speak("volume is more than 70 percent...please keep it below 70")
    else:
        volume.SetMasterVolumeLevel(currentVolumeDb + 3.0, None)
    # NOTE: -6.0 dB = half volume !


def set_vol(x):
    devices = AudioUtilities.GetSpeakers()
    interface = devices.Activate(
        IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
    volume = cast(interface, POINTER(IAudioEndpointVolume))

    # Get current volume
    currentVolumeDb = volume.GetMasterVolumeLevel()
    print(volume.GetMasterVolumeLevel())
    print(volume.GetVolumeRange())
    # volume.SetMasterVolumeLevel(currentVolumeDb  +6.0, None)
    # NOTE: -6.0 dB = half volume !


def decrese_vol():
    devices = AudioUtilities.GetSpeakers()
    interface = devices.Activate(
        IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
    volume = cast(interface, POINTER(IAudioEndpointVolume))

    # Get current volume
    currentVolumeDb = volume.GetMasterVolumeLevel()
    volume.SetMasterVolumeLevel(currentVolumeDb - 3.0, None)
    # NOTE: -6.0 dB = half volume !


# launch system apps
def app(query):
    # appdict = {"paint":,"excel":,"spreadsheet":,"word":,"powerpoint":,"presentation":,"google":,"chrome":}
    appname = query[7:]
    print(appname)
    speak('opening ' + appname)
    appname = appname.replace(" ", "")
    open(appname, match_closest=True)


# open websites
def website(query):
    site = query[5:]
    speak("opening website " + site)
    site = site.replace(" ", "")
    webbrowser.open("www." + site + ".com")


def kaggle(query):
    dset = query[14:]
    speak("these are the datasets")
    site = site.replace(" ", "")
    webbrowser.open("https://www.kaggle.com/search?q=" + dset)


def news(query):
    headlines = query[12:]
    speak("opening todays news" + headlines)
    headlines = headlines.replace(" ", "+")
    webbrowser.open("https://news.google.com/search?for=" + headlines + "&hl=en-IN&gl=IN&ceid=IN%3Aen")


def maps(query):
    location = query[9:]
    speak("finding" + location + "in google eearth")
    location = location.replace(" ", "+")
    webbrowser.open("https://earth.google.com/web/search/" + location)


# watch youtube videos
def youtube(query):
    utube = query[6:]
    speak("Opening youtube " + utube)
    pywhatkit.playonyt(utube)


def play_song(query):
    os.startfile(r"C:\Users\Arjun\OneDrive\Desktop\Spotify.lnk")
    for window in pyautogui.getAllWindows():
        if 'spotify' in window.title.lower():
            window.show()
            print('spotify window activated')
            song = query[21:]  # break down text list into single words for later usage
            time.sleep(5.5)
            pyautogui.click(x=72, y=107)  # this is the search bar location on my machine, when the window is maximized
            time.sleep(1)
            pyautogui.hotkey('ctrl', 'a')  # clearing the search bar
            pyautogui.press('backspace')  # clearing the search bar                time.sleep(0.5)
            pyautogui.write(song)  # because we assumed that the artist was the last word of the voice command
            time.sleep(3)
            pyautogui.click(x=766,
                            y=232)  # this is the play button location on my machine, when the window is maximized
            break


def google_search(ques):
    query = ques
    for j in search(query, tld="co.in", num=10, stop=10, pause=2):
        return (j)


def gogo():
    slepcom = ["Mention the duration you want me to go sleep",
               "Mention the duration you dont want me to disturb you",
               "Mention the duration you want me to be silent",
               "Mention the duration you want inner peace"]
    comm = random.choice(slepcom)
    print(comm)
    speak(comm)
    ss = listen()
    ss_int = int(ss)
    # ss = int(input())  #input is given manually we can convert it to be directly given by voice
    time.sleep(ss_int)
    print("Pikolo is back")
    speak("Pikolo is back...yyyyyaaaaaay")


def roast():
    rlist = ["Light travels faster than sound, which is why you seemed bright until you spoke.",
             "You have so many gaps in your teeth it looks like your tongue is in jail.",
             "time illa",
             "Your face makes onions cry.",
             " You bring everyone so much joy… when you leave the room."]
    rand = random.choice(rlist)
    print(rand)
    speak(rand)


def weather():
    city = "chennai"
    url = "https://www.google.com/search?q=" + "weather" + city
    html = requests.get(url).content
    soup = BeautifulSoup(html, 'html.parser')
    temp = soup.find('div', attrs={'class': 'BNeawe iBp4i AP7Wnd'}).text
    string = soup.find('div', attrs={'class': 'BNeawe tAd8D AP7Wnd'}).text
    data = string.split('\n')
    time = data[0]
    sky = data[1]
    listdiv = soup.findAll('div', attrs={'class': 'BNeawe s3v9rd AP7Wnd'})
    strd = listdiv[5].text
    pos = strd.find('Wind')
    other_data = strd[pos:]
    print("Temperature is", temp)
    print("Time: ", time)
    print('Sky Description:', sky)
    speak("Temperature is " + temp)
    speak("The sky is" + sky)


def wiki(query):
    speak("searching .......")
    query = query.replace("wikipedia", "")
    result = wikipedia.summary(query, sentences=2)
    speak("Accroding to wikipedia")
    speak(result)


def greet():
    hr = int(dt.datetime.now().hour)
    if (hr >= 0 and hr < 12):
        speak("Good morning !!")
    elif (hr >= 12 and hr < 18):
        speak("Good afternoon !!")
    elif (hr >= 18 and hr < 20):
        speak("Good evening  !!")
    speak("I am Pikolo, the Voice assistant, how can i help?")


def website(query):
    site = query[5:]
    speak("opening website " + site)
    site = site.replace(" ", "")
    webbrowser.open("www." + site + ".com")

def file(query):
	appname = query
	file_path = os.path.abspath(appname)
	print(file_path)
	
	os.system(file_path)

file("sample.txt")

    # appname = query[5:]
    # print(appname)
    # file <filename.format>
    # os.system()
    # os.path.abspath(appname)
    # for root, dirs, files in os.walk('.'):
    # if appname in files:
    # path=os.path.join(root, appname)
    # print(path)

    # return
    # else:
    # print("file not found")
    # return


def pikolo():
    schedule.run_pending()
    query = listen().lower()

    if "time" in query:
        now = dt.datetime.now().strftime("%d/%m/%Y %H:%M:%S")
        print(now)
        speak(now)

    elif "joke" in query:
        joke()

    elif "search dataset" in query:
        kaggle(query)

    elif "motivation" in query:
        motivation()

    elif "wikipedia" in query:
        wiki(query)

    elif "where is" in query:
        maps(query)

    elif ("headlines" in query) or ("news" in query):
        news(query)

    elif 'minimise the windows ' in query or 'minimise the window' in query:
        speak("minimize the window")
        pyautogui.hotkey('Win', 'd')

    elif 'maximize the windows' in query or 'maximize the window' in query:

        speak("maximizing the window")
        pyautogui.hotkey('Win', 'd')

    elif 'new tab' in query:
        pyautogui.hotkey('ctrl', 't')

    elif 'new file' in query:
        pyautogui.hotkey('ctrl', 'n')

    elif 'switch windows' in query or 'switch tab' in query:
        pyautogui.hotkey('ctrl', 'shift', 'tab')

    elif 'spotify' in query:
        play_song(query)


    elif 'switch the app' in query:
        pyautogui.hotkey('alt', 'tab')

    elif 'close window' in query:
        pyautogui.hotkey('alt', 'f4')

    elif "tell about you" in query:
        about()

    elif "who are you" in query:
        about()

    elif "increase volume" in query:
        increse_vol()

    elif "decrease volume" in query:
        decrese_vol()

    elif "what can you do" in query:
        func()
    elif "file" in query:
        file(query)

    elif "weather" in query:
        weather()

    elif "set" in query:
        y = query[10:]
        y = float(y)
        k = y / 100
        set_vol(k)

    # launch keyword to launch system apps
    elif "launch" in query:
        app(query)

    # open keyword to open websites
    elif "open" in query:
        website(query)

    # watch keyword to search and watch youtube videos
    elif "watch" in query:
        youtube(query)

    elif "search" in query:  # returns the best website for asked query

        result = query[6:]
        speak("giving best results for" + result)
        google_search(result)
        webbrowser.open(google_search(result))


    elif "how are you" in query:

        speak("i am fine human...how about you ?")

    elif query == "show commands":
        list()

    elif query == "sleep":
        gogo()

    elif (query == "who created you") or (query == "who made you"):
        devlopers()

    elif "roast" in query:
        roast()

    elif "remind" in query:
        remind()

    elif "speed" in query:
        print("working on it")
        speak("working on it")
        speed()

    elif query == "stop" or "end":
        exit()


cap = cv2.VideoCapture("vid (1).gif")
count = 0
commands = cv2.imread("list.png")
while True:
    ret, frame = cap.read()
    count += 1
    if count == cap.get(cv2.CAP_PROP_FRAME_COUNT):
        count = 0
        cap.set(cv2.CAP_PROP_POS_FRAMES, 0)
    cv2.putText(frame, "Pikolo-Voice Assistant", (250, 30), cv2.FONT_HERSHEY_PLAIN, 1.5, (255, 255, 255), 2)
    cv2.putText(frame, "Press S to speak and give commands", (150, 500), cv2.FONT_HERSHEY_PLAIN, 1.5, (255, 255, 255),
                2)
    cv2.putText(frame, "Press C to view commands", (220, 550), cv2.FONT_HERSHEY_PLAIN, 1.5, (255, 255, 255), 2)
    cv2.imshow("Pikolo", frame)
    key = cv2.waitKey(1)

    if key == ord('s'):
        pikolo()
    elif key == ord('c'):
        cv2.imshow("Commands", commands)
    if key & 0xFF == 27:
        break

cv2.destroyAllWindows()

