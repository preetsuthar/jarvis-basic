from ipaddress import summarize_address_range
import os
import logging
import subprocess
import ctypes
import pyttsx3
import requests
import datetime
import webbrowser
import pyjokes
import wikipedia
import smtplib
from email.message import EmailMessage
import imaplib
import email
import json
import psutil
from xmlrpc.server import DocXMLRPCServer
import speech_recognition as sr
from googletrans import Translator

class VirtualAssistant:
    def __init__(self):
        self.recognizer = sr.Recognizer()
        self.translator = Translator()
        self.engine = pyttsx3.init()
        self.setup_logging()
        self.greet_me()

    def setup_logging(self):
        logging.basicConfig(filename='assistant.log', level=logging.DEBUG,
                            format='%(asctime)s:%(levelname)s:%(message)s')

    def translate_text(self, text, target_language):
        try:
            translation = self.translator.translate(text, dest=target_language)
            self.speak(f"The translation is: {translation.text}")
        except Exception as e:
            logging.error(f"Error translating text: {e}")
            self.speak("Sorry, I couldn't translate the text.")

    def speak(self, text):
        self.engine.say(text)
        self.engine.runAndWait()

    def take_command(self):
        with sr.Microphone() as source:
            print("Listening...")
            self.recognizer.energy_threshold = 6000
            self.recognizer.pause_threshold = 0.5
            self.recognizer.operation_timeout = 5
            audio = self.recognizer.listen(source)

        try:
            print("Recognizing...")
            query = self.recognizer.recognize_google(audio, language='en-US')
            print(f"User said: {query}\n")
            return query.lower()
        except sr.UnknownValueError:
            self.speak("Sorry, I couldn't understand you. Can you please repeat?")
            return "None"
        except sr.RequestError:
            self.speak("Sorry, I'm having trouble connecting to the service. Please try again later.")
            return "None"
        except TimeoutError:
            self.speak("Sorry, I didn't hear anything. Please try again.")
            return "None"

    def greet_me(self):
        hour = datetime.datetime.now().hour
        if 0 <= hour < 12:
            self.speak("Good Morning!")
        elif 12 <= hour < 18:
            self.speak("Good Afternoon!")
        else:
            self.speak("Good Evening!")

    def open_website(self, url, site_name):
        self.speak(f"Opening {site_name}")
        webbrowser.open(url)

    def tell_joke(self):
        joke = pyjokes.get_joke()
        self.speak(joke)

    def wikipedia_search(self, query):
        try:
            results = wikipedia.summary(query, sentences=2)
            self.speak("According to Wikipedia")
            self.speak(results)
        except wikipedia.exceptions.DisambiguationError as e:
            logging.error(f"Wikipedia disambiguation error: {e}")
            self.speak("There are multiple results for this query, please be more specific.")
        except wikipedia.exceptions.PageError as e:
            logging.error(f"Wikipedia page error: {e}")
            self.speak("Sorry, I couldn't find any information on this topic.")
    
    def play_on_youtube(self, video):
        webbrowser.open(f"https://www.youtube.com/results?search_query={video}")

    def play_music(self):
        self.speak("Opening Spotify")
        spotify_path = "C:\\Users\\YourUsername\\AppData\\Roaming\\Spotify\\Spotify.exe"
        try:
            subprocess.Popen([spotify_path])
        except FileNotFoundError:
            self.speak("Spotify is not installed on your system.")

    def set_reminder(self, task, time):
        self.speak(f"Reminder set for {task} at {time}.")
        with open('reminders.txt', 'a') as file:
            file.write(f"{task} at {time}\n")
    
    def open_camera(self):
        subprocess.Popen(['start', 'microsoft.windows.camera:'], shell=True)

    def check_reminders(self):
        now = datetime.datetime.now().strftime("%H:%M")
        with open('reminders.txt', 'r') as file:
            reminders = file.readlines()

        new_reminders = []
        for reminder in reminders:
            task, reminder_time = reminder.strip().split(' at ')
            if now == reminder_time:
                self.speak(f"It's time to {task}.")
            else:
                new_reminders.append(reminder)

        with open('reminders.txt', 'w') as file:
            file.writelines(new_reminders)

    def send_email(self, subject, body, recipient):
        sender_email = 'your_email@gmail.com'  
        sender_password = 'your_email_password'  

        msg = EmailMessage()
        msg['Subject'] = subject
        msg['From'] = sender_email
        msg['To'] = recipient
        msg.set_content(body)

        try:
             with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
                    smtp.login(sender_email, sender_password)
                    smtp.send_message(msg)
                    self.speak("Email has been sent successfully.")
        except Exception as e:
                logging.error(f"Error sending email: {e}")
                self.speak("Sorry, I couldn't send the email.")

    def read_emails(self):
        try:
            self.speak("Connecting to email server.")
            mail = imaplib.IMAP4_SSL('imap.gmail.com')
            mail.login('your_email@gmail.com', 'your_email_password')
            mail.select('inbox')
            result, data = mail.search(None, 'UNSEEN')
            email_ids = data[0].split()
            if email_ids:
                latest_email_id = email_ids[-1]
                result, data = mail.fetch(latest_email_id, '(RFC822)')
                raw_email = data[0][1].decode('utf-8')
                email_message = email.message_from_string(raw_email)
                
                from_ = email_message['From']
                subject = email_message['Subject']
                self.speak(f"Email from {from_}, subject: {subject}")
                
                if email_message.is_multipart():
                    for part in email_message.walk():
                        if part.get_content_type() == "text/plain":
                            body = part.get_payload(decode=True).decode('utf-8')
                            self.speak(f"Email body: {body}")
                else:
                    body = email_message.get_payload(decode=True).decode('utf-8')
                    self.speak(f"Email body: {body}")
            else:
                self.speak("No new emails.")
        except Exception as e:
            logging.error(f"Error reading emails: {e}")
            self.speak("Sorry, I couldn't read the emails.")

    def read_document(self, file_path):
        try:
            doc = DocXMLRPCServer.Document(file_path)
            full_text = []
            for para in doc.paragraphs:
                full_text.append(para.text)
            return '\n'.join(full_text)
        except Exception as e:
            logging.error(f"Error reading document: {e}")
            self.speak("Sorry, I couldn't read the document.")
            return ""

    def summarize_text(self, text):
        try:
            summary = summarize_address_range(text)
            return summary
        except Exception as e:
            logging.error(f"Error summarizing text: {e}")
            self.speak("Sorry, I couldn't summarize the text.")
            return ""

    def summarize_document(self, file_path):
        self.speak(f"Reading the document {file_path}")
        text = self.read_document(file_path)
        if text:
            self.speak("Summarizing the document.")
            summary = self.summarize_text(text)
            self.speak("The summary is as follows:")
            self.speak(summary)

    def set_volume(self, level):
        self.speak(f"Setting volume to {level} percent.")
        os.system(f"nircmd.exe setsysvolume {int(level) * 65535 // 100}")

    def mute_volume(self):
        self.speak("Muting volume.")
        os.system("nircmd.exe mutesysvolume 1")

    def unmute_volume(self):
        self.speak("Unmuting volume.")
        os.system("nircmd.exe mutesysvolume 0")
        
    def open_calculator(self):
        subprocess.Popen(['calc.exe'])
        
    def open_cmd(self):
        subprocess.Popen(['cmd.exe'])
        
    def open_notepad(self):
        subprocess.Popen(['notepad.exe'])

    def open_discord(self):
        subprocess.Popen(r'C:\Users\preet\OneDrive\Desktop\Discord.lnk', shell=True)

    def set_brightness(self, level):
        self.speak(f"Setting brightness to {level} percent.")
        ctypes.windll.dxva2.SetMonitorBrightness(ctypes.windll.user32.MonitorFromWindow(ctypes.windll.user32.GetDesktopWindow(), 0), level)

    def find_nearby_places(self, place_type, location, radius=1000):
        api_key = 'your_google_maps_api_key'
        base_url = "https://maps.googleapis.com/maps/api/place/nearbysearch/json"
        params = {
            'location': location,
            'radius': radius,
            'type': place_type,
            'key': api_key
        }
        try:
            response = requests.get(base_url, params=params)
            results = response.json()

            if results['status'] == 'OK':
                places = results['results']
                self.speak(f"I found the following {place_type.replace('_', ' ')}s near you:")
                for place in places[:5]:
                    name = place['name']
                    vicinity = place['vicinity']
                    self.speak(f"{name}, located at {vicinity}")
            else:
                self.speak(f"Sorry, I couldn't find any {place_type.replace('_', ' ')}s near you.")
        except Exception as e:
            logging.error(f"Error fetching nearby places: {e}")
            self.speak("Sorry, I couldn't fetch the information right now.")
            
    def weather_update(self, city):
        api_key = 'your_openweather_api_key'
        base_url = f"http://api.openweathermap.org/data/2.5/weather?q={city}&appid={api_key}&units=metric"
        try:
            response = requests.get(base_url)
            weather_data = response.json()
            if weather_data['cod'] == 200:
                temp = weather_data['main']['temp']
                description = weather_data['weather'][0]['description']
                self.speak(f"The current temperature in {city} is {temp} degrees Celsius with {description}.")
            else:
                self.speak(f"Sorry, I couldn't get the weather information for {city}.")
        except Exception as e:
            logging.error(f"Error fetching weather data: {e}")
            self.speak("Sorry, I couldn't fetch the weather information right now.")
    
    def system_status(self):
        cpu_usage = psutil.cpu_percent(interval=1)
        memory_info = psutil.virtual_memory()
        self.speak(f"Current CPU usage is at {cpu_usage} percent.")
        self.speak(f"Current memory usage is at {memory_info.percent} percent.")

    def manage_tasks(self, task, action):
        if action == "add":
            with open('tasks.txt', 'a') as file:
                file.write(f"{task}\n")
            self.speak(f"Task '{task}' has been added.")
        elif action == "remove":
            with open('tasks.txt', 'r') as file:
                tasks = file.readlines()
            with open('tasks.txt', 'w') as file:
                tasks = [t for t in tasks if t.strip() != task]
                file.writelines(tasks)
            self.speak(f"Task '{task}' has been removed.")
        elif action == "list":
            self.speak("Here are your tasks:")
            with open('tasks.txt', 'r') as file:
                tasks = file.readlines()
                for t in tasks:
                    self.speak(t.strip())

    def open_chatgpt(self):
        self.speak("Opening ChatGPT in Chrome.")
        chrome_path = "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe"
        try:
            chatgpt_url = "https://www.example.com/chatgpt"  # Replace with actual ChatGPT URL
            webbrowser.get(chrome_path).open(chatgpt_url)
        except Exception as e:
            logging.error(f"Error opening ChatGPT in Chrome: {e}")
            self.speak("Sorry, I couldn't open ChatGPT in Chrome.")
            
    def run(self):
        self.speak("I am Jarvis, your virtual assistant. How can I help you today?")
        while True:
            self.check_reminders()
            query =self.take_command()
            if query == "None":
                continue
            if 'hello' in query:
                self.speak("Hello! How can I assist you?")
            elif 'Who are you' in query:
                self.speak("I am an artificial assistant created by preet sutharr and to help u make your work easy.")
            elif 'what can you do' in query:
                self.speak("I can do a variety of tasks, such as opening Google, opening YouTube, telling jokes, searching Wikipedia, playing music, setting reminders, controlling system volume and brightness, reading your emails, managing tasks, checking weather updates, and more. Just ask!")
            elif 'bye' in query:
                self.speak("Goodbye! Have a great day!")
                break
            elif 'google' in query:
                self.open_website("https://www.google.com", "Google")
            elif 'youtube' in query:
                self.open_website("https://www.youtube.com", "YouTube")
            elif 'joke' in query:
                self.tell_joke()
            elif 'wikipedia' in query:
                self.speak("What should I search on Wikipedia?")
                search_query = self.take_command()
                self.wikipedia_search(search_query)
            elif 'play music' in query:
                self.play_music()
            elif 'open discord' in query:
                self.open_discord()
            elif 'open command prompt' in query or 'open cmd' in query:
                self.open_cmd()
            elif 'open camera' in query:
                self.open_camera()
            elif 'open calculator' in query:
                self.open_calculator()
            elif 'set reminder' in query:
                self.speak("What do you want to be reminded about?")
                task = self.take_command()
                self.speak("At what time?")
                time = self.take_command()
                self.set_reminder(task, time)
            elif 'send email' in query:
                self.speak("What is the subject of the email?")
                subject = self.take_command()
                self.speak("What should I say in the email?")
                body = self.take_command()
                self.speak("To whom should I send the email?")
                recipient = self.take_command() 
                self.send_email(subject, body, recipient)
            elif 'read email' in query:
                self.read_emails()
            elif 'set volume' in query:
                self.speak("What volume level should I set? (0 to 100)")
                level = int(self.take_command())
                self.set_volume(level)
            elif 'mute volume' in query:
                self.mute_volume()
            elif 'unmute volume' in query:
                self.unmute_volume()
            elif 'set brightness' in query:
                self.speak("What brightness level should I set? (0 to 100)")
                level = int(self.take_command())
                self.set_brightness(level)
            elif 'summarize document' in query:
                self.speak("Which document would you like me to summarize?")
                document_name = self.take_command()
                self.summarize_document(document_name)
            elif 'find nearby' in query:
                self.speak("What type of place are you looking for?")
                place_type = self.take_command()
                self.speak("Please provide your location in 'latitude,longitude' format.")
                location = self.take_command()
                self.find_nearby_places(place_type, location)
            elif 'search on google' in query:
                self.speak("What would you like to search on Google?")
                search_query = self.take_command()
                google_search_url = "https://www.google.com/search?q=" + search_query.replace(' ', '+')
                self.open_website(google_search_url, "Google Search")
            elif 'weather' in query:
                self.speak("Which city's weather would you like to know?")
                city = self.take_command()
                self.weather_update(city)
            elif 'system status' in query:
                self.system_status()
            elif 'add task' in query:
                self.speak("What task would you like to add?")
                task = self.take_command()
                self.manage_tasks(task, "add")
            elif 'remove task' in query:
                self.speak("What task would you like to remove?")
                task = self.take_command()
                self.manage_tasks(task, "remove")
            elif 'list tasks' in query:
                self.manage_tasks("", "list")
            elif 'open chatgpt' in query or 'open chat gpt' in query:
                self.open_chatgpt()
if __name__ == "__main__":
    assistant = VirtualAssistant()
    assistant.run()