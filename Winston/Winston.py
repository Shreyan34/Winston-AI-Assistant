#importing all the ncessary modules for our ai
from win32com.client import Dispatch #this module shall help us to make our speak function
import datetime #this module shall help to get the time
import speech_recognition as sr #this module shall help our module to get commands from the user
import wikipedia #for getting search results from Wikipedia, we import the wikipedia module
import webbrowser #this module shall help us to open an url on the user's request

#the following three modules shall help us to track the phone numbers
import phonenumbers 
from phonenumbers import geocoder
from phonenumbers import carrier


def speak(audio):
    '''
    This function will speak the string parameter 'audio'
    '''
    speak = Dispatch(("SAPI.SpVoice"))
    speak.Speak(audio)

def wishMe():
    '''
    This function helps our AI to greet the user and give its introduction
    '''
    hour = int(datetime.datetime.now().hour)
    if hour >= 0 and hour < 12:
        speak("Good morning!")
    elif hour>=12 and hour < 18:
        speak("Good afternoon!")
    else:
        speak("Good evening!")
    
    speak("I am Winston. Please tell me how can I help you?")

def takeCommand():
    '''
    This is the function which helps our AI to take commands from the user
    '''
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening...")
        r.pause_threshold = 1
        audio = r.listen(source)
    
    try:
        print("Recognizing...")
        query = r.recognize_google(audio, language='en-in')
        print(f'User said: {query}\n')

    except Exception as e:
        # print(e)
        speak("Can you please say that again?")
        return "None"
    return query

def main():
    '''
    This is the main method of our program
    '''
    wishMe()

    query = takeCommand().lower()

    if 'wikipedia' in query:
        speak("Searching Wikipedia...")
        query = query.replace("wikipedia", "")
        results = wikipedia.summary(query, sentences=2)
        speak("According to Wikipedia")
        #print(results)
        speak(results)

    elif 'open youtube' in query:
        speak("Opening YouTube")
        webbrowser.open('youtube.com')

    elif 'open google' in query:
        speak("Opening Google")
        webbrowser.open('google.com')
    
    elif 'time' in query:
        strTime = datetime.datetime.now().strftime("%H:%M:%S")
        speak(f'The time is {strTime}')
    
    elif 'track a number' or 'track the number' or 'track number' in query:
        speak("Enter the number you want to track")
        number = input("Enter the number you want to track: ")

        ch_number = phonenumbers.parse(number, "CH")
        country = geocoder.description_for_number(ch_number, "en")
        print(f"The country is: {country}")    
        
        service_number = phonenumbers.parse(number, "RO")
        provider = carrier.name_for_number(service_number, "en")
        # print(provider)
        print(f"And the service provider is: {provider}")

    elif 'alarm' in query:
        speak("Enter the time")
        time = input(": Enter the time :")
        
        while True:
           time_ac = datetime.datetime.now()
           now = time_ac.strftime("%H:%M:%S")
           if now == time:
               speak("Time to wake up!")
           elif now > time:
               break

main()