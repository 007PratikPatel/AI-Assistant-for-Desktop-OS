import speech_recognition as sr
import pyttsx3
import pywhatkit as kit
import datetime
import wikipedia
import pyjokes
import smtplib
import os
import config
import subprocess

engine = pyttsx3.init() 


def speak(audio):
    engine.say(audio)
    engine.runAndWait()


def takeCommand():

    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("J.A.R.V.I.S is Listening...")
        speak("jarvis is listening")
        r.pause_threshold = 1.7
        audio = r.listen(source)

    try:
        query = r.recognize_google(audio, language='en-in')
        print(query)

    except Exception as e:
        speak('Please say the command again.')
        return "None"

    return query


def sendEmail(to, content):
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.ehlo()
    server.starttls()

    # Enable low security in gmail
    server.login(config.EMAIL_ADDRESS, config.PASSWORD)
    server.sendmail('animen313@gmail.com', to, content)
    server.close()


def run_jarvis():

    query = takeCommand().lower()

    if 'introduce yourself' in query:
        os.startfile(
            "F:\\IMCA\\IMCA 2nd SEM\\Project-1\\JARVIS.mp3")

    elif 'play' in query:
        play = query.replace('play', '')
        speak('playing'+play)
        kit.playonyt(play)

    elif 'search' in query:
        search = query.replace('search', '')
        speak('searching'+search)
        kit.search(search)

    elif 'what is' in query or 'what are' in query:
        search = query.replace('what is' or 'what are',
                               'what is' or 'what are')
        speak('searching'+search)
        kit.search(search)

    elif 'how' in query or 'how to' in query:
        search = query.replace('how' or 'how to', 'how' or 'how to')
        speak('searching'+search)
        kit.search(search)

    elif 'tell me about' in query:
        query = query.replace('tell me abot', '')
        info = wikipedia.summary(query, sentences=5)
        print(info)
        speak("According to Wikipedia"+info)

# Opening some files

    elif 'presentation' in query:
        speak('opening presentation')
        os.startfile(
            'C:\\Pratik\\Project-1\\AI Virtual Assistant.pptx')

    elif 'project report' in query:
        speak('opening project report')
        os.startfile(
            'C:\\Pratik\\Project-1\\AI Virtual Assistant.docx')

    elif 'my code' in query:
        speak('opening my code')
        os.startfile(
            'C:\\Pratik\\Project-1\\J.A.R.V.I.S.py')

    elif 'project workspace' in query or 'project work space' in query:
        speak('opening project workspace')
        os.startfile(
            'C:\\Pratik\\Project-1')

    elif 'avengers endgame' in query:
        speak('opening')
        os.startfile("D:\\Entertainment\\Hollywood\\Avengers.Endgame.mp4")

# Opening Application

    elif 'zoom' in query:
        speak('opening zoom')
        subprocess.Popen(
            ["C:\\Users\\Pratik\\AppData\\Roaming\\Zoom\\bin\\Zoom.exe"])

    elif 'file explorer' in query:
        speak('opening file explorer')
        subprocess.Popen(["C:\\Windows\\explorer.exe"])

    elif 'notepad' in query:
        speak('opening notepad')
        subprocess.Popen(["C:\\Windows\\notepad.exe"])

    elif 'vs code' in query:
        speak('opening vs code')
        subprocess.Popen(
            ["C:\\Users\\Pratik\\AppData\\Local\\Programs\\Microsoft VS Code\\Code.exe"])

    elif 'excel' in query:
        speak('opening excel')
        subprocess.Popen(
            ["C:\\Program Files (x86)\\Microsoft Office\\Office16\\EXCEL.EXE"])

    elif 'task manager' in query:
        speak('opening task manager')
        subprocess.Popen(
            ["C:\\Windows\\System32\\Taskmgr.exe"])

    elif 'powerpoint' in query:
        speak('opening powerpoint')
        subprocess.Popen(
            ["C:\Program Files (x86)\Microsoft Office\Office16\POWERPNT.EXE"])

    elif 'word' in query:
        speak('opening word')
        subprocess.Popen(
            ["C:\Program Files (x86)\Microsoft Office\Office16\WINWORD.EXE"])

    elif 'calculator' in query:
        speak('opening calculator')
        subprocess.Popen(
            ["C:\Windows\System32\calc.exe"])

    elif 'virtualbox' in query:
        speak('opening VirtualBox')
        subprocess.Popen(
            ["C:\Program Files\Oracle\VirtualBox\VirtualBox.exe"])

    elif 'xampp' in query:
        speak('opening xampp')
        subprocess.Popen(
            ["C:\\xampp\\xampp-control.exe"])

    elif 'chrome' in query:
        speak('opening chrome')
        subprocess.Popen(
            ["C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"])

# Opening Web Apps

    elif 'youtube' in query:
        speak('opening youtube')
        subprocess.Popen(
            ['C:\Program Files (x86)\Google\Chrome\Application\chrome.exe', 'www.youtube.com'])

    elif 'google keep' in query:
        speak('opening google keep')
        subprocess.Popen(
            ['C:\Program Files (x86)\Google\Chrome\Application\chrome.exe', 'keep.google.com'])

    elif 'google drive' in query:
        speak('opening google drive')
        subprocess.Popen(
            ['C:\Program Files (x86)\Google\Chrome\Application\chrome.exe', 'drive.google.com'])

    elif 'whatsapp' in query:
        speak('opening whatsapp')
        subprocess.Popen(
            ['C:\Program Files (x86)\Google\Chrome\Application\chrome.exe', 'web.whatsapp.com'])

    elif 'google photos' in query:
        speak('opening google photos')
        subprocess.Popen(
            ['C:\Program Files (x86)\Google\Chrome\Application\chrome.exe', 'photos.google.com/u/1/?pli=1'])

    elif 'google doc' in query:
        speak('opening google docs')
        subprocess.Popen(
            ['C:\Program Files (x86)\Google\Chrome\Application\chrome.exe', 'docs.google.com/document/u/1/?tgif=d'])

    elif 'google spreadsheet' in query:
        speak('google spreadsheets')
        subprocess.Popen(
            ['C:\Program Files (x86)\Google\Chrome\Application\chrome.exe', 'docs.google.com/spreadsheets/u/1/?tgif=d'])

    elif 'google form' in query:
        speak('opening google forms')
        subprocess.Popen(
            ['C:\Program Files (x86)\Google\Chrome\Application\chrome.exe', 'docs.google.com/forms/u/1/?tgif=d'])

    elif 'google slide' in query:
        speak('opening google slides')
        subprocess.Popen(
            ['C:\Program Files (x86)\Google\Chrome\Application\chrome.exe', 'docs.google.com/presentation/u/1/?tgif=d'])

    elif 'google meet' in query:
        speak('opening google meet')
        subprocess.Popen(
            ['C:\Program Files (x86)\Google\Chrome\Application\chrome.exe', 'meet.google.com'])

    elif 'calendar' in query:
        speak('opening calendar')
        subprocess.Popen(
            ['C:\Program Files (x86)\Google\Chrome\Application\chrome.exe', 'calendar.google.com/calendar/u/0/r/month'])

    elif 'gmail' in query:
        speak('opening gmail')
        subprocess.Popen(
            ['C:\Program Files (x86)\Google\Chrome\Application\chrome.exe', 'mail.google.com/mail/u/1/#inbox'])

    elif 'google map' in query:
        speak('opening google map')
        subprocess.Popen(
            ['C:\Program Files (x86)\Google\Chrome\Application\chrome.exe', 'maps.google.com'])

    elif 'google classroom' in query:
        speak('opening google classroom')
        subprocess.Popen(
            ['C:\Program Files (x86)\Google\Chrome\Application\chrome.exe', 'classroom.google.com/u/1/h'])

    elif 'bookmyshow' in query:
        speak('opening bookmyshow')
        subprocess.Popen(
            ['C:\Program Files (x86)\Google\Chrome\Application\chrome.exe', 'in.bookmyshow.com'])

    elif 'flipkart' in query:
        speak('opening flipkart')
        subprocess.Popen(
            ['C:\Program Files (x86)\Google\Chrome\Application\chrome.exe', 'www.flipkart.com'])

    elif 'amazon' in query:
        speak('opening amazon')
        subprocess.Popen(
            ['C:\Program Files (x86)\Google\Chrome\Application\chrome.exe', 'www.amazon.in'])

    elif 'spotify' in query:
        speak('opening spotify')
        subprocess.Popen(
            ['C:\Program Files (x86)\Google\Chrome\Application\chrome.exe', 'open.spotify.com'])

    elif 'linkedin' in query:
        speak('opening linkedin')
        subprocess.Popen(
            ['C:\Program Files (x86)\Google\Chrome\Application\chrome.exe', 'www.linkedin.com'])

    elif 'twitter' in query:
        speak('opening twitter')
        subprocess.Popen(
            ['C:\Program Files (x86)\Google\Chrome\Application\chrome.exe', 'twitter.com/home'])

    elif 'instagram' in query:
        speak('opening instagram')
        subprocess.Popen(
            ['C:\Program Files (x86)\Google\Chrome\Application\chrome.exe', 'www.instagram.com'])

    elif 'hotstar' in query:
        speak('opening hotstar')
        subprocess.Popen(
            ['C:\Program Files (x86)\Google\Chrome\Application\chrome.exe', 'www.hotstar.com/in'])

    elif 'pu mis' in query:
        speak('opening pu mis')
        subprocess.Popen(
            ['C:\Program Files (x86)\Google\Chrome\Application\chrome.exe', 'pumis.in'])

    elif 'time' in query:
        time = datetime.datetime.now().strftime('%I:%M %p')
        print(time)
        speak('Current time is'+time)

    elif 'date' in query:
        date = datetime.datetime.now().strftime('%A, %d %b, %Y')
        print(date)
        speak('Today is'+date)

    elif 'joke' in query:
        query = query.replace('joke', '')
        jokes = pyjokes.get_joke()
        print(jokes)
        speak(jokes)

#Sending Emails

    elif 'pratik' in query:
        try:
            speak("What should I say?")
            content = takeCommand()
            to = "pratikbpatel3250@gmail.com"
            sendEmail(to, content)
            speak("Email has been sent!")
        except Exception as e:
            print(e)
            speak("Sorry, I am not able to send this email")

    elif 'prateek' in query:
        try:
            speak("What should I say?")
            content = takeCommand()
            to = "pratikbpatel007@gmail.com"
            sendEmail(to, content)
            speak("Email has been sent!")
        except Exception as e:
            print(e)
            speak("Sorry, I am not able to send this email")

    elif 'anjali mam' in query:
        try:
            speak("What should I say?")
            content = takeCommand()
            to = "anjali.mahavar42078@paruluniversity.ac.in"
            sendEmail(to, content)
            speak("Email has been sent!")
        except Exception as e:
            print(e)
            speak("Sorry, I am not able to send this email")

    elif 'shin-chan' in query or 'sinchan' in query:
        try:
            speak("What should I say?")
            content = takeCommand()
            to = "snchnnhr@gmail.com"
            sendEmail(to, content)
            speak("Email has been sent!")
        except Exception as e:
            print(e)
            speak("Sorry, I am not able to send this email")

    elif 'goku' in query:
        try:
            speak("What should I say?")
            content = takeCommand()
            to = "patelgoku007@gmail.com"
            sendEmail(to, content)
            speak("Email has been sent!")
        except Exception as e:
            print(e)
            speak("Sorry, I am not able to send this email")

    elif 'hulk' in query:
        try:
            speak("What should I say?")
            content = takeCommand()
            to = "007hulkpatel@gmail.com"
            sendEmail(to, content)
            speak("Email has been sent!")
        except Exception as e:
            print(e)
            speak("Sorry, I am not able to send this email")


run_jarvis()
