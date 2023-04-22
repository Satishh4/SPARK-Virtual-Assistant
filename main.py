import random
import speech_recognition as sr

recognizer = sr.Recognizer()
import pyttsx3
import wikipedia
import webbrowser
import time
import wolframalpha
import pyjokes
import os
import smtplib
import imaplib
import requests
import email
from bs4 import BeautifulSoup
import datetime
import subprocess
import pytesseract
from PIL import Image

# tkinter GUI begin


# tkinter GUI end


voiceEngine = pyttsx3.init()
voices = voiceEngine.getProperty('voices')

art = """ 
 .d8888b.  8888888b.         d8888 8888888b.  888    d8P         d888        .d8888b.  
d88P  Y88b 888   Y88b       d88888 888   Y88b 888   d8P         d8888       d88P  Y88b 
Y88b.      888    888      d88P888 888    888 888  d8P            888       888    888 
 "Y888b.   888   d88P     d88P 888 888   d88P 888d88K             888       888    888 
    "Y88b. 8888888P"     d88P  888 8888888P"  8888888b            888       888    888 
      "888 888          d88P   888 888 T88b   888  Y88b           888       888    888 
Y88b  d88P 888         d8888888888 888  T88b  888   Y88b          888   d8b Y88b  d88P 
 "Y8888P"  888        d88P     888 888   T88b 888    Y88b       8888888 Y8P  "Y8888P"  

"""

art2 = """ 

 _    ______ _____________     _______          ____________________________________________   _______
  \  /   |  |_____/  |   |     |_____||         |_____|_____|______ |  |______  |   |_____| \  |  |   
   \/  __|__|    \_  |   |_____|     ||_____    |     ______________|________|  |   |     |  \_|  |   

 Created by [Satish] and [Jayanth] under Guidance of [
  _____                  _              __      _______ 
 |  __ \           /\   (_)             \ \    / / ____|
 | |  | |_ __     /  \   _  __ _ _   _   \ \  / / |  __ 
 | |  | | '__|   / /\ \ | |/ _` | | | |   \ \/ /| | |_ |
 | |__| | |_    / ____ \| | (_| | |_| |    \  / | |__| |
 |_____/|_(_)  /_/    \_\ |\__,_|\__, |     \/   \_____|
                       _/ |       __/ |                 
                      |__/       |___/                  
] from Civil Department as a final year project, 

 SPARK 1.0 is a powerful virtual assistant that allows you to access your system and perform various operations using just your voice commands.
Speak to SPARK 1.0 and it will help you with your daily tasks like sending emails, scheduling meetings, playing music, searching the web and much more. You can also customize SPARK 1.0 as per your needs and preferences.
Get started with SPARK 1.0 now by saying "Hey SPARK, how can you help me?


say "  HELP" to get some example commands
"""

if len(voices) >= 2:
    voiceEngine.setProperty('voice', voices[1].id)
else:
    print(art)
    print(art2)

asname = ""

engine = pyttsx3.init()


def speak(text):
    engine.say(text)
    engine.runAndWait()


def wish():
    print(art)
    print(art2)
    print("Wishing.")
    time = int(datetime.datetime.now().hour)
    global asname
    if time >= 0 and time < 12:
        speak("Good Morning!")
    elif time < 18:
        speak("Good Afternoon!")
    else:
        speak("Good Evening!")
    asname = "Spark 1 point o"
    speak("I am your Virtual Assistant from SVIT,")
    speak(asname)
    print("I am your Virtual Assistant,", asname)


def getName():
    speak("Welcome! How can I help you today?")


def takeCommand():
    with sr.Microphone() as source:
        print("I am Listening...")
        recognizer.pause_threshold = 1
        audio = recognizer.listen(source)
    try:
        print("Recognizing the command")
        command = recognizer.recognize_google(audio, language='en-in')
        print(f"Command is: {command}\n")

        if 'search' in command.lower():
            if 'desktop' in command.lower():
                query = command.lower().replace('search', '').replace('desktop', '').strip()
                directory = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')
                results = [file for file in os.listdir(directory) if query in file.lower()]
                print(f"Search results for '{query}' on desktop: {results}")
            elif 'documents' in command.lower():
                query = command.lower().replace('search', '').replace('documents', '').strip()
                directory = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Documents')
                results = [file for file in os.listdir(directory) if query in file.lower()]
                print(f"Search results for '{query}' in documents: {results}")
    except Exception as e:
        print(e)
        print("Unable to Recognize your voice.")
        return "None"

    return command


def openApp(appName):
    try:
        # Replace the path with the path of the executable file of the app you want to open
        path = "C:\\Program Files\\" + appName + ".exe"
        subprocess.Popen(path)
        speak(f"Opening {appName}")
    except Exception as e:
        print(str(e))
        speak("Unable to open the app")


def searchFile(fileName):
    try:
        # Replace the path with the path of the folder where you want to search for the file
        path = "C:\\Users\\Username\\Documents\\"
        for root, dirs, files in os.walk(path):
            if fileName in files:
                filePath = os.path.join(root, fileName)
                os.startfile(filePath)
                speak(f"Opening {fileName}")
                return
        speak(f"File {fileName} not found")
    except Exception as e:
        print(str(e))
        speak("Unable to search for the file")


# Function to read a doc file and convert it to speech
def readDocFile():
    # Change the path to the path of the doc file on your desktop
    path = os.path.join(os.path.expanduser('~'), 'Desktop', 'example.docx')
    text = ''
    try:
        # Extract text from the doc file using pytesseract OCR
        text = pytesseract.image_to_string(path)
    except Exception as e:
        print(str(e))
    if text:
        # Convert the extracted text to speech using pyttsx3
        engine.say(text)
        engine.runAndWait()
    else:
        print('Unable to extract text from the file.')


CONTACTS = {
    'Satish': 'sathishkumars.20cs@saividya.ac.in',
    'Jayanth': 'sjayanth.19cs@saividya.ac.in',
    'Kishan': 'kishankumarm.19ec@saividya.ac.in',
    'Abhishek': 'abhishekm.18cs@saividya.ac.in',
    'Mohan': 'gtmohankumar.19cs@saividya.ac.in'
}


def sendEmail(to_name):
    to_email = CONTACTS.get(to_name)
    if to_email is None:
        print(f"Error: {to_name} not found in contacts")
        return
    print(f"Sending mail to {to_name} ({to_email})")
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.ehlo()
    server.starttls()
    # paste your email id and password in the respective places
    server.login('sathishkumars.20cs@saividya.ac.in', '9066488570')
    content = input("What should I say in the email? ")
    server.sendmail('sathishkumars.20cs@saividya.ac.in', to_email, content)
    server.close()


# Example usage:
def sendEmailUsage():
    names = ['Satish', 'Jayanth', 'Kishan', 'Abhishek', 'Mohan']
    for name in names:
        sendEmail(name)


def getWeather(city_name):
    # base url from where we extract weather report
    baseUrl = "http://api.openweathermap.org/data/2.5/weather?"
    url = baseUrl + "appid=" + 'd850f7f52bf19300a9eb4b0aa6b80f0d' + "&q=" + city_name
    response = requests.get(url)
    x = response.json()

    # If there is no error, getting all the weather conditions
    if x["cod"] != "404":
        y = x["main"]
        temp = y["temp"]
        temp -= 273
        pressure = y["pressure"]
        humidity = y["humidity"]
        desc = x["weather"]
        description = desc[0]["description"]
        info = (" Temperature= " + str(temp) + "°C" + "\n atmospheric pressure (hPa) =" + str(pressure) +
                "\n humidity = " + str(humidity) + "%" + "\n description = " + str(description))
        print(info)
        speak("Here is the weather report at")
        speak(city_name)
        speak(info)
    else:
        speak(" City Not Found ")



def getNews():
    url = 'https://www.bbc.com/'
    response = requests.get(url)
    soup = BeautifulSoup(response.content, 'html.parser')
    headlines = []
    for headline in soup.find_all('h3', class_='gs-c-promo-heading__title gel-paragon-bold nw-o-link-split__text'):
        headlines.append(headline.text)
    return headlines


# Call the function and print the returned list of headlines
headlines = getNews()
for headline in headlines:
    print(headline)

if __name__ == '__main__':

    uname = ''
    asname = ''
    os.system('cls')
    wish()
    getName()
    print(uname)

    while True:

        command = takeCommand().lower()
        print(command)

        if "jarvis" in command:
            wish()

        elif 'how are you' in command:
            speak("I am fine, Thank you")
            speak("How are you, ")
            speak(uname)

        elif "what time is it" in command:
            now = datetime.datetime.now()
            current_time = now.strftime("%H:%M:%S")
            speak(f"The current time is {current_time}")


        elif "good morning" in command or "good afternoon" in command or "good evening" in command:
            speak("A very" + command)
            speak("Thank you for wishing me! Hope you are doing well!")

        elif 'fine' in command or "good" in command:
            speak("It's good to know that your fine")

        elif "who are you" in command:
            speak("I am your virtual assistant.")

        elif "change my name to" in command:
            speak("What would you like me to call you,")
            uname = takeCommand()
            speak('Hello again,')
            speak(uname)

        elif "change name" in command:
            speak("What would you like to call me,")
            assname = takeCommand()
            speak("Thank you for naming me!")


        elif "what's your name" in command:
            speak("People call me")
            speak(assname)

        elif 'help' in command:
            speak('Here are some things you can ask me to do:')
            speak('--------------------------------------')
            speak('Command                                | Action')
            speak('--------------------------------------')
            speak('Open Google                            | Opens Google in your default web browser')
            speak('Play music                             | Plays a random song from your music directory')
            speak('Joke                                   | Tells you a random joke')
            speak('Turn off computer                      | Shuts down your computer after confirmation')
            speak('Play [video/song] on YouTube           | Searches YouTube for a video or song and plays it')
            speak('Mail                                   | Sends an email to a recipient')
            speak('Read document                          | Reads a document on your desktop using OCR')
            speak('Read mail                              | Reads an email in your inbox')
            speak('How are you                            | Asks how I am doing')
            speak("What time is it                        | Tells you the current time")
            speak("Good morning/afternoon/evening         | Greets you based on the time of day")
            speak("Who are you                            | Tells you who I am")
            speak("Change my name to [name]               | Changes your name that I use to address you")
            speak("Change name                            | Changes my name")
            speak("What's your name                       | Tells you my name")


        elif 'read document' in command:
            speak('Please provide the name of the document on your desktop')
            document_name = takeCommand()
            document_path = os.path.join(os.path.expanduser("~"), "Desktop", document_name)
            try:
                with Image.open(document_path) as img:
                    text = pytesseract.image_to_string(img)
                    speak(text)
            except Exception as e:
                print(e)
                speak('Sorry, I was unable to read the document')

        elif 'read mail' in command:
            speak('Please provide the name of the person whose mail you want to read')
            to_name = takeCommand().title()
            to_email = CONTACTS.get(to_name)
            if to_email is None:
                speak(f"Error: {to_name} not found in contacts")
            else:
                try:
                    server = imaplib.IMAP4_SSL('imap.gmail.com', 993)
                    server.login('your email id', 'your email password')
                    server.select('inbox')
                    result, data = server.search(None, 'ALL')
                    email_ids = data[0].split()
                    for email_id in email_ids:
                        result, data = server.fetch(email_id, '(RFC822)')
                        raw_email = data[0][1]
                        email_message = email.message_from_bytes(raw_email)
                        if email_message['To'] == to_email:
                            speak(f"From: {email_message['From']}")
                            speak(f"Subject: {email_message['Subject']}")
                            speak(f"Body: {email_message.get_payload()}")
                            break
                except Exception as e:
                    print(e)
                    speak('Sorry, I was unable to read the email')


        elif 'time' in command:
            strTime = datetime.datetime.now()
            curTime = str(strTime.hour) + "hours" + str(strTime.minute) + \
                      "minutes" + str(strTime.second) + "seconds"
            speak(uname)
            speak(f" the time is {curTime}")
            print(curTime)

        elif 'wikipedia' in command:
            speak('Searching Wikipedia')
            command = command.replace("wikipedia", "")
            results = wikipedia.summary(command, sentences=3)
            speak("These are the results from Wikipedia")
            print(results)
            speak(results)

        elif 'open youtube' in command:
            speak("Here you go, the Youtube is opening\n")
            print("""
  ,,,,,,,,,,,,,,,,,,,,,,,    ..   ...              .......       ...            
 ,,,,,,,,,,,,,,,,,,,,,,,,,   ...  ..                 ..          ...            
 ,,,,,,,,,,@@,,,,,,,,,,,,,    .. ..  .....  ...  ..  ..  ... ... .......  ......
 ,,,,,,,,,,@@@@@@@,,,,,,,,     .... ...  .. ...  ..  ..  ... ... ...  ..  ..  ..
 ,,,,,,,,,,@@,,,,,,,,,,,,,     ...  ...  .. ...  ..  ..  ... ... ...  .. #......
 ,,,,,,,,,,,,,,,,,,,,,,,,,     #..  ... #.. ...  ..  ..  ... ... ...  ..  ..  ..
  ,,,,,,,,,,,,,,,,,,,,,,,      #..    ...    ... ..  ..   ../ .. ... ..    ...                                                      

                                                                              """)
            webbrowser.open("youtube.com")

        elif 'open google' in command:
            speak("Opening Google\n")
            webbrowser.open("google.com")


        elif 'play music' in command or "play song" in command:

            speak("Enjoy the music!")

            music_dirs = ['/home/user/music', '/mnt/music', '/media/music']

            all_songs = []

            for music_dir in music_dirs:

                if os.path.isdir(music_dir):
                    songs = os.listdir(music_dir)

                    all_songs.extend([os.path.join(music_dir, song) for song in songs])

            if not all_songs:

                speak("No music files found.")

            else:

                song_path = random.choice(all_songs)

                os.startfile(song_path)

        elif 'joke' in command:
            speak(pyjokes.get_joke())

        elif 'calculate' in command:
            # extract the mathematical expression from the user command
            expression = command.replace('calculate', '').strip()

            # evaluate the expression using the built-in eval function
            try:
                result = eval(expression)
                speak(f"The result of {expression} is {result}")
            except Exception as e:
                speak(f"Sorry, I couldn't calculate {expression}. Error: {e}")

        elif 'turn off computer' in command:
            speak("Are you sure you want to turn off the computer?")

            # listen for confirmation
            confirmation = command()

            if 'yes' in confirmation:
                os.system("shutdown /s /t 1")
            else:
                speak("Then why the hell did you say turn off.")

        elif 'play' in command and ('video' in command or 'music' in command) and 'on youtube' in command:
            # extract the search query from the user command
            query = command.replace('play', '').replace('video', '').replace('music', '').replace('on youtube',
                                                                                                  '').strip()

            # perform a YouTube search using the query
            search_url = f'https://www.youtube.com/results?search_query={query}'
            html = requests.get(search_url).content
            soup = BeautifulSoup(html, 'html.parser')
            video_url = f"https://www.youtube.com/watch?v={soup.find_all('a')[3].get('href').split('=')[1]}"

            # play the video or music using the webbrowser module
            webbrowser.open(video_url)


        elif 'mail' in command:
            try:
                speak("Whom should I send the mail")
                to = input()
                speak("What is the body?")
                content = takeCommand()
                sendEmail(to, content)
                speak("Email has been sent successfully !")
            except Exception as e:
                print(e)
                speak("I am sorry, not able to send this email")

        elif 'exit' in command:
            speak("Thanks for giving me your time")
            exit()

        elif "will you be my gf" in command or "will you be my bf" in command:
            speak("I'm not sure about that, may be you should give me some time")

        elif "i love you" in command:
            speak("Thank you! But, It's a pleasure to hear it from you.")

        elif "weather" in command:
            speak(" Please tell your city name ")
            print("City name : ")
            cityName = takeCommand()
            getWeather(cityName)


        elif "what is" in command or "who is" in command:
            client = wolframalpha.Client("5UV4LG-U5VELG94Q2")
            res = client.query(command)
            try:
                print(next(res.results).text)
                speak(next(res.results).text)
            except StopIteration:
                print("No results")


        elif 'search' in command:
            command = command.replace("search", "")
            webbrowser.open(command)

        elif 'news' in command:
            getNews()

        elif "don't listen" in command or "stop listening" in command:
            speak("for how much time you want to stop me from listening commands")
            a = int(takeCommand())
            time.sleep(a)
            print(a)

        elif 'shutdown system' in command:
            speak("Hold On a Sec ! Your system is on its way to shut down")
            subprocess.call('shutdown / p /f')

        elif "restart" in command:
            subprocess.call(["shutdown", "/r"])

        elif "sleep" in command:
            speak("Setting in sleep mode")
            subprocess.call("shutdown / h")

        elif "write a note" in command:
            speak("What should i write, sir")
            note = takeCommand()
            file = open('spark.txt', 'w')
            speak("Sir, Should i include date and time")
            snfm = takeCommand()
            if 'yes' in snfm or 'sure' in snfm:
                strTime = datetime.datetime.now().strftime("% H:% M:% S")
                file.write(strTime)
                file.write(" :- ")
                file.write(note)
            else:
                file.write(note)
        else:
            speak("Sorry, I am not able to understand you")

