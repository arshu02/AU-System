import speech_recognition as sr
import pyttsx3
import pyautogui
import os
import subprocess
import psutil
import webbrowser
import wikipedia
import datetime
import json
import keyboard
import requests
from time import sleep
import random
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import winreg
import win32com.client
from datetime import datetime, timedelta
import threading
import re

class VoiceAssistant:
    def __init__(self):
        # Initialize speech recognition
        self.recognizer = sr.Recognizer()
        
        # Initialize text-to-speech engine
        self.engine = pyttsx3.init()
        self.engine.setProperty('rate', 150)
        
        # Get available voices
        voices = self.engine.getProperty('voices')
        # Set a female voice if available
        for voice in voices:
            if "female" in voice.name.lower():
                self.engine.setProperty('voice', voice.id)
                break
        
        # Store active reminders
        self.reminders = []
        
        # Load commands configuration
        self.commands = {
            'open': self.open_application,
            'search': self.search_web,
            'screenshot': self.take_screenshot,
            'system': self.system_info,
            'volume': self.control_volume,
            'close': self.close_application,
            'type': self.type_text,
            'wikipedia': self.search_wikipedia,
            'weather': self.get_weather,
            'time': self.get_time,
            'date': self.get_date,
            'email': self.send_email,
            'remind': self.set_reminder,
            'music': self.control_music,
            'joke': self.tell_joke,
            'news': self.get_news,
            'translate': self.translate_text,
            'calculate': self.calculate,
            'schedule': self.manage_schedule
        }

        # Greeting variations
        self.greetings = [
            "Hello! How can I help you today?",
            "Hi there! What can I do for you?",
            "Hey! I'm ready to assist you.",
            "Greetings! How may I help you?"
        ]

        # Common applications paths
        self.app_paths = {
            "chrome": "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe",
            "firefox": "C:\\Program Files\\Mozilla Firefox\\firefox.exe",
            "word": "WINWORD.EXE",
            "excel": "EXCEL.EXE",
            "powerpoint": "POWERPNT.EXE",
            "notepad": "notepad.exe",
            "calculator": "calc.exe",
            "paint": "mspaint.exe",
            "cmd": "cmd.exe",
            "explorer": "explorer.exe"
        }

    def speak(self, text):
        """Convert text to speech with more natural pauses"""
        print(f"Assistant: {text}")
        sentences = text.split('.')
        for sentence in sentences:
            if sentence.strip():
                self.engine.say(sentence.strip())
                self.engine.runAndWait()
                sleep(0.3)  # Natural pause between sentences

    def listen(self):
        """Listen for voice input with improved error handling"""
        with sr.Microphone() as source:
            print("Listening...")
            self.recognizer.adjust_for_ambient_noise(source, duration=1)
            try:
                audio = self.recognizer.listen(source, timeout=5, phrase_time_limit=10)
                text = self.recognizer.recognize_google(audio)
                print(f"You said: {text}")
                return text.lower()
            except sr.UnknownValueError:
                self.speak("I didn't quite catch that. Could you please repeat?")
                return None
            except sr.RequestError:
                self.speak("I'm having trouble connecting to the speech recognition service.")
                return None
            except Exception as e:
                self.speak("Something went wrong. Please try again.")
                print(f"Error: {str(e)}")
                return None

    def open_application(self, app_name):
        """Open an application with enhanced app support"""
        try:
            app_name = app_name.lower()
            
            # Handle web browsers and search queries
            if "google" in app_name or "chrome" in app_name:
                if os.path.exists(self.app_paths["chrome"]):
                    os.startfile(self.app_paths["chrome"])
                    self.speak("Opening Google Chrome")
                else:
                    webbrowser.open("https://www.google.com")
                    self.speak("Opening Google in default browser")
                return
            elif "firefox" in app_name:
                if os.path.exists(self.app_paths["firefox"]):
                    os.startfile(self.app_paths["firefox"])
                    self.speak("Opening Firefox")
                else:
                    webbrowser.open("https://www.mozilla.org/firefox")
                    self.speak("Opening Firefox website")
                return
            
            # Handle other applications
            if app_name in self.app_paths:
                if app_name in ["word", "excel", "powerpoint"]:
                    # Use Win32COM for Office applications
                    win32com.client.Dispatch(f"Applications.{app_name}")
                else:
                    os.startfile(self.app_paths[app_name])
                self.speak(f"Opening {app_name}")
            else:
                # Try to find the application in Start Menu
                try:
                    subprocess.Popen(f"start {app_name}", shell=True)
                    self.speak(f"Opening {app_name}")
                except:
                    self.speak(f"I couldn't find {app_name}. Would you like me to search for it online?")
                    
        except Exception as e:
            self.speak(f"I encountered an error while trying to open {app_name}")
            print(f"Error: {str(e)}")

    def search_web(self, query):
        """Search the web using default browser"""
        try:
            # Clean up the query
            query = query.strip()
            
            # Handle special search cases
            if query.startswith("for "):
                query = query[4:]  # Remove "for " from the beginning
            
            # Encode the query for URL
            encoded_query = requests.utils.quote(query)
            
            # Open in default browser
            url = f"https://www.google.com/search?q={encoded_query}"
            webbrowser.open_new_tab(url)
            self.speak(f"Searching for {query}")
        except Exception as e:
            self.speak("I had trouble performing the web search")
            print(f"Error: {str(e)}")

    def take_screenshot(self, filename=None):
        """Take a screenshot"""
        if not filename:
            filename = f"screenshot_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.png"
        pyautogui.screenshot(filename)
        self.speak(f"Screenshot saved as {filename}")

    def system_info(self):
        """Get system information"""
        cpu = psutil.cpu_percent()
        memory = psutil.virtual_memory().percent
        info = f"CPU usage is {cpu}%. Memory usage is {memory}%"
        self.speak(info)

    def control_volume(self, action):
        """Control system volume"""
        if action == "up":
            pyautogui.press("volumeup", presses=5)
        elif action == "down":
            pyautogui.press("volumedown", presses=5)
        elif action == "mute":
            pyautogui.press("volumemute")
        self.speak(f"Volume {action}")

    def close_application(self, app_name):
        """Close an application"""
        try:
            os.system(f"taskkill /f /im {app_name}.exe")
            self.speak(f"Closed {app_name}")
        except:
            self.speak(f"Couldn't close {app_name}")

    def type_text(self, text):
        """Type text"""
        pyautogui.write(text)
        self.speak(f"Typed: {text}")

    def search_wikipedia(self, query):
        """Search Wikipedia"""
        try:
            result = wikipedia.summary(query, sentences=2)
            self.speak(result)
        except:
            self.speak("Sorry, I couldn't find that on Wikipedia")

    def get_weather(self, city):
        """Get weather information"""
        try:
            # You'll need to sign up for a free API key at OpenWeatherMap
            api_key = "YOUR_API_KEY"  # Replace with your API key
            url = f"http://api.openweathermap.org/data/2.5/weather?q={city}&appid={api_key}&units=metric"
            response = requests.get(url)
            data = response.json()
            
            if response.status_code == 200:
                temp = data['main']['temp']
                desc = data['weather'][0]['description']
                humidity = data['main']['humidity']
                self.speak(f"In {city}, it's {temp}°C with {desc}. The humidity is {humidity}%")
            else:
                self.speak("I couldn't fetch the weather information at the moment")
        except Exception as e:
            self.speak("Sorry, I couldn't get the weather information")
            print(f"Error: {str(e)}")

    def get_time(self):
        """Get current time"""
        current_time = datetime.now().strftime("%I:%M %p")
        self.speak(f"The current time is {current_time}")

    def get_date(self):
        """Get current date"""
        current_date = datetime.now().strftime("%B %d, %Y")
        self.speak(f"Today is {current_date}")

    def send_email(self, command):
        """Send email using voice commands"""
        try:
            # Extract email details from command
            # Format: "email to recipient@example.com subject Hello message This is the message"
            parts = command.split()
            if len(parts) < 6:
                self.speak("Please provide recipient, subject, and message for the email")
                return
            
            recipient = parts[parts.index("to") + 1]
            subject_index = parts.index("subject")
            message_index = parts.index("message")
            
            subject = " ".join(parts[subject_index + 1:message_index])
            message = " ".join(parts[message_index + 1:])
            
            # Email configuration
            sender_email = "your_email@gmail.com"  # Replace with your email
            sender_password = "your_app_password"   # Replace with your app password
            
            # Create message
            msg = MIMEMultipart()
            msg['From'] = sender_email
            msg['To'] = recipient
            msg['Subject'] = subject
            msg.attach(MIMEText(message, 'plain'))
            
            # Send email
            with smtplib.SMTP('smtp.gmail.com', 587) as server:
                server.starttls()
                server.login(sender_email, sender_password)
                server.send_message(msg)
            
            self.speak("Email sent successfully")
        except Exception as e:
            self.speak("Sorry, I couldn't send the email")
            print(f"Error: {str(e)}")

    def control_music(self, command):
        """Control music playback"""
        try:
            if "play" in command:
                keyboard.press_and_release('play/pause media')
                self.speak("Playing music")
            elif "pause" in command:
                keyboard.press_and_release('play/pause media')
                self.speak("Pausing music")
            elif "next" in command:
                keyboard.press_and_release('next track')
                self.speak("Playing next track")
            elif "previous" in command:
                keyboard.press_and_release('previous track')
                self.speak("Playing previous track")
        except Exception as e:
            self.speak("I couldn't control the music playback")
            print(f"Error: {str(e)}")

    def tell_joke(self):
        """Tell a random joke"""
        try:
            response = requests.get("https://official-joke-api.appspot.com/random_joke")
            if response.status_code == 200:
                joke_data = response.json()
                setup = joke_data['setup']
                punchline = joke_data['punchline']
                self.speak(f"{setup}")
                sleep(1)
                self.speak(f"{punchline}")
            else:
                self.speak("Sorry, I couldn't fetch a joke right now")
        except Exception as e:
            self.speak("Sorry, I couldn't tell a joke right now")
            print(f"Error: {str(e)}")

    def get_news(self):
        """Get latest news headlines"""
        try:
            # You'll need to sign up for a free API key at NewsAPI
            api_key = "YOUR_API_KEY"  # Replace with your API key
            url = f"https://newsapi.org/v2/top-headlines?country=us&apiKey={api_key}"
            response = requests.get(url)
            
            if response.status_code == 200:
                news_data = response.json()
                articles = news_data['articles'][:3]  # Get top 3 headlines
                self.speak("Here are the top headlines:")
                for i, article in enumerate(articles, 1):
                    self.speak(f"{i}. {article['title']}")
            else:
                self.speak("I couldn't fetch the news at the moment")
        except Exception as e:
            self.speak("Sorry, I couldn't get the news")
            print(f"Error: {str(e)}")

    def translate_text(self, command):
        """Translate text to specified language"""
        try:
            # Format: "translate Hello to Spanish"
            text = command.split("translate")[1].split("to")[0].strip()
            target_lang = command.split("to")[1].strip()
            
            # You'll need to sign up for a free API key at Google Cloud
            url = f"https://translation.googleapis.com/language/translate/v2?key=YOUR_API_KEY"
            payload = {
                "q": text,
                "target": target_lang
            }
            response = requests.post(url, json=payload)
            
            if response.status_code == 200:
                translation = response.json()['data']['translations'][0]['translatedText']
                self.speak(f"The translation is: {translation}")
            else:
                self.speak("I couldn't translate the text at the moment")
        except Exception as e:
            self.speak("Sorry, I couldn't translate the text")
            print(f"Error: {str(e)}")

    def calculate(self, expression):
        """Perform basic calculations"""
        try:
            # Remove dangerous functions and limit to basic operations
            safe_dict = {'add': '+', 'plus': '+', 'minus': '-', 'subtract': '-', 
                        'multiply': '*', 'times': '*', 'divided': '/', 'divide': '/'}
            for key, value in safe_dict.items():
                expression = expression.replace(key, value)
            result = eval(expression)
            self.speak(f"The result is {result}")
        except Exception as e:
            self.speak("I couldn't perform that calculation")
            print(f"Error: {str(e)}")

    def manage_schedule(self, command):
        """Manage calendar schedule"""
        try:
            outlook = win32com.client.Dispatch("Outlook.Application")
            calendar = outlook.GetNamespace("MAPI").GetDefaultFolder(9)
            
            if "add" in command:
                # Format: "schedule add meeting tomorrow at 2pm for 1 hour"
                # Parse date and time from command
                appointment = outlook.CreateItem(1)
                appointment.Subject = "Meeting"
                appointment.Start = datetime.now() + timedelta(days=1)  # Example
                appointment.Duration = 60  # minutes
                appointment.Save()
                self.speak("Event added to calendar")
            elif "check" in command:
                items = calendar.Items
                items.Sort("[Start]")
                items.IncludeRecurrences = "True"
                
                today = datetime.now()
                tomorrow = today + timedelta(days=1)
                
                self.speak("Here are your upcoming events:")
                for item in items:
                    if today <= item.Start <= tomorrow:
                        self.speak(f"{item.Subject} at {item.Start.strftime('%I:%M %p')}")
        except Exception as e:
            self.speak("I couldn't manage the schedule")
            print(f"Error: {str(e)}")

    def set_reminder(self, command):
        """Set a reminder with time and message"""
        try:
            # Parse the reminder command
            # Format: "remind me in 5 minutes to take a break" or
            # Format: "remind me at 3:30 pm to call John"
            
            if "in" in command:
                # Handle relative time
                match = re.search(r'in (\d+) (minute|minutes|hour|hours)', command)
                if match:
                    amount = int(match.group(1))
                    unit = match.group(2)
                    
                    # Calculate reminder time
                    if 'hour' in unit:
                        reminder_time = datetime.now() + timedelta(hours=amount)
                    else:
                        reminder_time = datetime.now() + timedelta(minutes=amount)
                    
                    # Extract message (everything after "to")
                    message = command.split(" to ")[-1]
                    
                    # Create and start reminder thread
                    thread = threading.Thread(
                        target=self._reminder_thread,
                        args=(reminder_time, message)
                    )
                    thread.daemon = True
                    thread.start()
                    
                    # Store reminder
                    self.reminders.append({
                        'time': reminder_time,
                        'message': message,
                        'thread': thread
                    })
                    
                    self.speak(f"I'll remind you {message} in {amount} {unit}")
                else:
                    self.speak("I couldn't understand the reminder time format")
            
            elif "at" in command:
                # Handle absolute time
                match = re.search(r'at (\d{1,2}):?(\d{2})?\s*(am|pm)?', command)
                if match:
                    hour = int(match.group(1))
                    minute = int(match.group(2)) if match.group(2) else 0
                    meridiem = match.group(3)
                    
                    # Convert to 24-hour format if needed
                    if meridiem:
                        if meridiem.lower() == 'pm' and hour != 12:
                            hour += 12
                        elif meridiem.lower() == 'am' and hour == 12:
                            hour = 0
                    
                    # Create reminder time for today
                    now = datetime.now()
                    reminder_time = now.replace(hour=hour, minute=minute, second=0, microsecond=0)
                    
                    # If the time is already passed, set it for tomorrow
                    if reminder_time < now:
                        reminder_time += timedelta(days=1)
                    
                    # Extract message (everything after "to")
                    message = command.split(" to ")[-1]
                    
                    # Create and start reminder thread
                    thread = threading.Thread(
                        target=self._reminder_thread,
                        args=(reminder_time, message)
                    )
                    thread.daemon = True
                    thread.start()
                    
                    # Store reminder
                    self.reminders.append({
                        'time': reminder_time,
                        'message': message,
                        'thread': thread
                    })
                    
                    self.speak(f"I'll remind you to {message} at {reminder_time.strftime('%I:%M %p')}")
                else:
                    self.speak("I couldn't understand the reminder time format")
            else:
                self.speak("Please specify when you want to be reminded using 'in' or 'at'")
        
        except Exception as e:
            self.speak("I had trouble setting the reminder")
            print(f"Error setting reminder: {str(e)}")
    
    def _reminder_thread(self, reminder_time, message):
        """Background thread for handling a reminder"""
        try:
            # Calculate sleep duration
            sleep_duration = (reminder_time - datetime.now()).total_seconds()
            if sleep_duration > 0:
                sleep(sleep_duration)
                
                # When time is up, speak the reminder
                self.speak(f"Reminder: {message}")
                
                # Remove the reminder from the list
                self.reminders = [r for r in self.reminders 
                                if r['time'] != reminder_time 
                                and r['message'] != message]
        
        except Exception as e:
            print(f"Error in reminder thread: {str(e)}")

    def process_command(self, command):
        """Process voice command with enhanced natural language understanding"""
        if not command:
            return True

        # Handle greetings
        greetings = ["hello", "hi", "hey", "good morning", "good afternoon", "good evening"]
        if any(greeting in command for greeting in greetings):
            self.speak(random.choice(self.greetings))
            return True

        # Handle gratitude
        if "thank you" in command or "thanks" in command:
            self.speak("You're welcome! Is there anything else I can help you with?")
            return True

        # Handle goodbye
        if "goodbye" in command or "bye" in command or "stop" in command or "exit" in command:
            self.speak("Goodbye! Have a great day!")
            return False

        # Handle combined search commands
        if "search for" in command:
            query = command.split("search for", 1)[1].strip()
            self.search_web(query)
            return True
        
        if "google" in command and "search" in command:
            # Extract the search query
            if "for" in command:
                query = command.split("for", 1)[1].strip()
            else:
                query = command.replace("google", "").replace("search", "").strip()
            self.search_web(query)
            return True

        # Process regular commands
        words = command.split()
        action = words[0]
        params = ' '.join(words[1:]) if len(words) > 1 else ""

        if action in self.commands:
            self.commands[action](params)
        else:
            # Try to understand context
            if "what" in command and "time" in command:
                self.get_time()
            elif "what" in command and "date" in command:
                self.get_date()
            elif "how" in command and "weather" in command:
                city = command.split("in")[-1].strip() if "in" in command else "local"
                self.get_weather(city)
            elif "tell" in command and "joke" in command:
                self.tell_joke()
            else:
                # If command not understood, try web search
                self.speak("I'll search the web for that information")
                self.search_web(command)
        
        return True

    def run(self):
        """Main loop with improved interaction"""
        current_hour = datetime.now().hour
        if current_hour < 12:
            greeting = "Good morning!"
        elif current_hour < 18:
            greeting = "Good afternoon!"
        else:
            greeting = "Good evening!"
            
        self.speak(f"{greeting} I'm your voice assistant. How can I help you today?")
        
        while True:
            command = self.listen()
            if not self.process_command(command):
                break

if __name__ == "__main__":
    assistant = VoiceAssistant()
    assistant.run() 