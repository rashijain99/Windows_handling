# Artificial intelligence project - Windows Handling using Speech Recognition
# Importing necessary modules

import pyttsx3                                    
import datetime                                   
import speech_recognition as sr                    
import wikipedia                                  
import webbrowser                                 
import random                                     
import os                                                                                 
import subprocess as sp   
import requests
import pywhatkit as kit
import smtplib     
import wolframalpha
import ctypes
import winshell
import time
import imdb
import screen_brightness_control as sbc

###################################################
engine = pyttsx3.init()                           
voices = engine.getProperty('voices')             
engine.setProperty('voice',voices[1].id)          
###################################################

folder=[]
dirpath=[]    


def speak(audio):
    engine.say(audio)
    engine.runAndWait()
    
def wishMe():
    hour = int(datetime.datetime.now().hour)
    if hour>=0 and hour<12:
        speak("Good Morning!")
    elif hour>=12 and hour<18:
        speak("Good Afternoon")
    else:
        speak("Good Evening")
    speak(" Mam I am your desktop assistant JARVIS.\
          Tell me how may I help you?") 

def takeCommand():
    r = sr.Recognizer()
    with sr.Microphone() as source :
        r.adjust_for_ambient_noise(source)
        print("LISTENING...")
        r.pause_threshold = 0.5
        audio=r.listen(source)
    try:
        print("Recognizing...")
        query = r.recognize_google(audio, language='en-in')
        print("You Said :",query)
    except Exception as e:
        print(e)
        print("COULD NOT RECOGNIZE SAY AGAIN PLEASE..")
        return "None"
    return query

def tellDay(): 
    day = datetime.datetime.today().weekday() + 1
    Day_dict = {1: 'Monday', 2: 'Tuesday',  
                3: 'Wednesday', 4: 'Thursday',  
                5: 'Friday', 6: 'Saturday', 
                7: 'Sunday'}
    if day in Day_dict.keys(): 
        day_of_the_week = Day_dict[day] 
        print(day_of_the_week) 
        speak("The day is " + day_of_the_week) 
  
def tellTime(): 
    time = str(datetime.datetime.now()) 
    print(time) 
    hour = time[11:13] 
    min = time[14:16] 
    speak( "The time is " + hour + "Hours and" + min + "Minutes")     
    
def open_notepad():
    programName= "notepad.exe"
    sp.Popen([programName])
    
def close_notepad():
    sp.call('TASKKILL /IM notepad.exe /F') 
    
def open_calculator():
    programName= "calc.exe"
    sp.Popen([programName])
    
def close_calculator():
    sp.call('TASKKILL /IM calculator.exe /F')   
    
def open_paint():
    programName = "mspaint.exe"
    sp.Popen([programName])   
    
def close_paint():
    sp.call('TASKKILL /IM mspaint.exe /F')    
    
def open_camera():
    sp.run('start microsoft.windows.camera:', shell=True)        
 
def open_cmd():
    os.system('start cmd')  
    
def open_Excel():
    os.startfile('EXCEL.EXE') 
    
def close_Excel():
    sp.call('TASKKILL /IM EXCEL.EXE /F')
    
def open_powershell():
    os.startfile("powershell.exe")  
    
def close_powershell():
    sp.call('TASKKILL /IM powershell.exe /F') 
    
def open_Powerpoint():
    os.startfile('POWERPNT.EXE') 
    
def close_Powerpoint():
    sp.call('TASKKILL /IM POWERPNT.EXE /F')    
 
def open_vscode():
    os.startfile(r"C:\Users\user\AppData\Local\Programs\Microsoft VS Code\Code.exe")
    
def close_vscode():
    sp.call('TASKKILL /IM Code.exe /F') 
    
def search_on_wikipedia(query):
    results = wikipedia.summary(query, sentences=2)
    return results

def play_on_youtube(video):
    kit.playonyt(video)

def search_on_google(query):
    kit.search(query)
    
def send_whatsapp_message(number, message):
    kit.sendwhatmsg_instantly(f"+91{number}", message)

def get_random_joke():
    headers = {
        'Accept': 'application/json'
    }
    res = requests.get("https://icanhazdadjoke.com/", headers=headers).json()
    return res["joke"]

def get_random_advice():
    res = requests.get("https://api.adviceslip.com/advice").json()
    return res['slip']['advice']
    
def search_movie():
    moviesdb = imdb.IMDb()
    speak('which movie would you like to search')
    text = takeCommand()

    movies = moviesdb.search_movie(text)
  
    speak("Searching for " + text)
    if len(movies) == 0:
        speak("No result found")
    else:
  
        speak("I found these:")
        for movie in movies:
            title = movie['title']
            year = movie['year']
            
            speak(f'{title}-{year}')
            print(f'{title}-{year}')
            
            info = movie.getID()
            movie = moviesdb.get_movie(info)
            title = movie['title']
            year = movie['year']
            rating = movie['rating']
            plot = movie['plot outline']
            
            if year < int(datetime.datetime.now().strftime("%Y")):
                speak(
                    f'{title}was released in {year} has IMDB rating of {rating}.\
                    The plot summary of movie is{plot}')
                print(
                    f'{title}was released in {year} has IMDB rating of {rating}.\
                    The plot summary of movie is{plot}')    
                break
  
            else:
                speak(
                    f'{title}will release in {year} has IMDB rating of {rating}.\
                    The plot summary of movie is{plot}')
                print(
                    f'{title}will release in {year} has IMDB rating of {rating}.\
                    The plot summary of movie is{plot}')    
                break
            
        time.sleep(5)       

def goodbye():
    hour2=int(datetime.datetime.now().hour)
    if hour2>=0 and hour2<12:
        speak("Thank You! Have a Nice Day MAM")
    elif hour2>=12 and hour2<18:
        speak("Thank You! Have a Good Day MAM")
    else:
        speak("Thank You! Have a Good Night MAM")
        
###############################################################

if __name__ == "__main__":
   wishMe()
   while True:
       query = takeCommand().lower()
       
       if 'what is my name' in query or "my name" in query:
           speak("your name is Rashi")
           
       elif 'what is your name' in query or "your name" in query:
           speak("My name is Jarvis, Mam!")
           
       elif 'how are you' in query:
           speak(" I am fine Mam, thank you!!")
           
       elif 'hey jarvis' in query  or 'hi jarvis' in query:
           speak(" I AM AWAKE, Mam")
           
       elif 'thank you jarvis' in query or "thank you" in query:
           speak("Your Welcome Mam")
           
       elif 'what can you do' in query:
           speak("Anything You Want, Mam")
           
       elif "who i am" in query:
            speak("If you talk then definitely your human.")
    
       elif "why you came to world" in query or "why you come" in query:
            speak("Thanks to Rashi. further It's a secret")
 
       elif "who are you" in query:
            speak("I am your virtual assistant.")    
            
       elif 'temperature outside' in query:
            speak("Let me check, Mam")
            webbrowser.open("https://www.google.com/search?sxsrf=ALeKk008XudKSCgyx_YBG_dZR4wotlbJHA%3A1603470699933&source=hp&ei=awWTX-bfNorUz7sPpJmBSA&q=temperature&oq=tem&gs_lcp=CgZwc3ktYWIQAxgAMgwIIxAnEJ0CEEYQgAIyBAgAEEMyBwgAELEDEEMyAggAMgUIABCxAzIFCC4QsQMyCAgAELEDEIMBMggIABCxAxCDATICCAAyAggAOgcIIxDqAhAnOgQIIxAnOgcIIxDJAxAnOgoIABCxAxCDARBDOgoILhDHARCvARAnOggIABDJAxCRAjoFCAAQkQJQnBVY_xhgriFoAXAAeACAAbQDiAHbBpIBBzItMi4wLjGYAQCgAQGqAQdnd3Mtd2l6sAEK&sclient=psy-ab")
            break
    
       elif "tell day" in query or "tell de" in query:                                    
            tellDay() 
            
       elif "tell time" in query or "time" in query: 
            tellTime()   
             
       elif 'C drive' in query or "open C drive" in query or "C Drive" in query:
            speak("opening c drive")
            os.startfile("C:")
           
       elif 'D drive' in query or "open D drive" in query or "D Drive" in query:
            speak("opening d drive")
            os.startfile("D:")
            
       elif 'E drive' in query or "open E drive" in query or "E Drive" in query:
            speak("opening e drive")
            os.startfile("E:")   
            
       elif "open task manager" in query:
            os.startfile(r"C:\Windows\System32\Taskmgr.exe")
            
       elif "open windows explorer" in query:
            os.startfile(r"C:\Windows\explorer.exe")   
            
       elif "open whatsapp" in query:
           os.startfile(r'C:\Users\user\AppData\Local\WhatsApp\app-2.2226.5\WhatsApp.exe')
           break
        
       elif 'wikipedia' in query:
              speak('What do you want to search on Wikipedia, Mam?')
              search_query = takeCommand().lower()
              results = search_on_wikipedia(search_query)
              speak(f"According to Wikipedia, {results}")
              speak("For your convenience, I am printing it on the screen Mam.")
              print(results)
              
       elif 'open youtube' in query:
           speak("opening youtube")
           webbrowser.open("www.youtube.com")
           
       elif 'open google' in query:
           speak("opening google")
           webbrowser.open("www.google.com")
           
       elif 'open facebook' in query:
           speak("searching facebook")
           webbrowser.open("https://www.facebook.com/")
           
       elif 'open amazon' in query:
           speak("Searching amazon")
           webbrowser.open("https://www.amazon.in/")
        
       elif 'whatsapp browser' in query:
           webbrowser.open("https://web.whatsapp.com/")     
           
       elif 'open spotify' in query:
            os.system('start spotify:')  
            
       elif 'open command prompt' in query or 'open cmd' in query:
            open_cmd()
            
       elif 'open camera' in query:
            open_camera()  
           
       elif 'open notepad' in query:
           speak("opening notepad")
           open_notepad()
           
       elif 'close notepad' in query:                    
           speak("closing notepad")
           close_notepad()    
           
       elif 'open paint' in query:
           speak("opening paint")
           open_paint()

       elif 'close paint' in query:
           speak("closing paint")
           close_paint()    
           
       elif 'open calculator' in query:
           speak("opening calculator")
           open_calculator()
           
       elif 'close calculator' in query:
           speak("closing calculator")
           close_calculator() 
           
       elif 'open shell' in query:
           speak("Opening shell")
           open_powershell()
           
       elif 'close shell' in query:
           speak("closing shell")
           close_powershell()    
           
       elif 'Excel' in query or 'open excel' in query:
           speak("opening excel")
           open_Excel()
            
       elif 'close excel' in query:
           speak("closing excel")
           close_Excel()     
          
       elif 'open powerpoint presentation' in query or "open ppt" in query or "open powerpoint" in query:
            speak("opening Power Point presentation")
            open_Powerpoint()
             
       elif 'close powerpoint presentation' in query or "close ppt" in query or "close powerpoint" in query:
            speak("closing Power Point presentation")
            close_Powerpoint()
    
       elif 'open code' in query or 'open vs code' in query or 'open visual studio code' in query:
           speak("opening vs code")
           open_vscode()

       elif 'close code' in query or 'close vs code' in query or 'close visual studio code' in query:
           speak("closing vs code")
           close_vscode()    
           
       elif 'locate' in query or 'search place' in query:
            speak("What do you want to locate, Mam")
            location = takeCommand()
            speak("Hold On Sir I Will Show You Where"+location+"is")
            webbrowser.open("https://www.google.com/maps/place/"+location+"")   
            break
           
       elif 'play music' in query or 'play song' in query:
            speak('What do you want to play on Youtube, Mam?')
            video = takeCommand()
            play_on_youtube(video)
            break

       elif 'search on google' in query:
            speak('What do you want to search on Google, Mam?')
            query = takeCommand()
            search_on_google(query)
            break

       elif "send whatsapp message" in query:
            speak('On what number should I send the message Mam? Please enter in the console: ')
            number = input("Enter the number: ")
            speak("What is the message Mam?")
            message = takeCommand()
            send_whatsapp_message(number, message)
            speak("I've sent the message Mam.")
            break
        
       elif 'joke' in query or 'tell me a joke' in query:
            speak(f"Hope you like this one Mam")
            joke = get_random_joke()
            speak(joke)
            speak("For your convenience, I am printing it on the screen Mam.")
            print(joke)

       elif "advice" in query or 'give me a advice' in query:
            speak(f"Here's an advice for you, Mam")
            advice = get_random_advice()
            speak(advice)
            speak("For your convenience, I am printing it on the screen Mam.")
            print(advice) 
            
       elif 'search movie' in query:
           search_movie() 
           
       elif 'change background' in query:
            ctypes.windll.user32.SystemParametersInfoW(20,
                                                       0,
                                                       "Location of wallpaper",
                                                       0)
            speak("Background changed successfully")    
       
       elif "brightness" in query or "change brightness" in query:
            a = sbc.get_brightness()
            speak(f"current brightness is , {a}")
            print(a)
            speak("what you want to do? Enter choice")
            print("1. set brightness from the current brightness to 100%")
            print("2. set brightness from the current brightness to 75%")
            print("3. set brightness from the current brightness to 50%")
            print("4. set brightness from the current brightness to 25%")
            print("5. set brightness from the current brightness to 0%")
            
            m = int(takeCommand())
            if m == 1:
               sbc.set_brightness(100)
               speak("brightness set to 100%")
            elif m == 2 :
               sbc.set_brightness(75)
               speak("brightness set to 75%")
            elif m == 3:
               sbc.set_brightness(50)
               speak("brightness set to 50%") 
            elif m == 4:
               sbc.set_brightness(25)
               speak("brightness set to 25%") 
            elif m == 5:
               sbc.set_brightness(0)
               speak("brightness set to 0%")  
               
               
       elif 'empty recycle bin' in query or 'recycle bin' in query:
            winshell.recycle_bin().empty(confirm = False, show_progress = False, sound = True)
            speak("Recycle Bin Recycled")
    
       elif "don't listen" in query or "stop listening" in query:
            speak("for how much time you want to stop jarvis from listening commands")
            speak("please give time in seconds!")
            a = int(takeCommand())
            speak("ok time start now!")
            time.sleep(a)
            print(a)
            speak("back again!")
       
       elif "write a note" in query or "write note" in query:
           speak("What should i write, Mam")
           note = takeCommand()
           file = open('jarvis.txt', 'w')
           file.write(note)
        
       elif "show note" in query:
           speak("Showing Notes")
           file = open("jarvis.txt", "r")
           print(file.read())
           speak(file.read(6))
           
       elif 'game' in query or "play game" in query:
           speak('select the game - 1 Rock paper scissors game! , 2 number gussing game!')
           speak('Please select a number ')
           print("select the game - 1 Rock paper scissors game! , 2 number gussing game!")
           n = int(takeCommand())
           if n == 1:
               power = r"D:\PROJECTS - python\rock_paper_scissors.py"
               os.startfile(power)
               break 
           elif n == 2 :
               power = r"D:\PROJECTS - python\number_gussing.py"
               os.startfile(power)
               break 
           
       elif 'lock window' in query:
                speak("locking the device")
                ctypes.windll.user32.LockWorkStation()
                break
        
       elif "restart" in query:
           sp.call(["shutdown", "/r"])
            
       elif "hibernate" in query or "sleep" in query:
           speak("Hibernating")
           sp.call("shutdown / h")

       elif "log off" in query or "sign out" in query:
           speak("Make sure all the application are closed before sign-out")
           time.sleep(5)
           sp.call(["shutdown", "/l"])       
              
       elif 'shutdown system' in query:
           speak("Hold On a Sec ! Your system is on its way to shut down")
           sp.call('shutdown / p /f')
           break
       
       elif 'bye' in query or 'quit' in query:
           goodbye()
           break    
       
###------------------------------------------------------------------------###

           
           
           
           
          
         
           
           
   
        