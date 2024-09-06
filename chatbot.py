import datetime
import urllib.parse
import urllib.request
import speech_recognition as sr
import win32com.client
import webbrowser
import wikipedia
import os
import wave
import threading
import pyaudio as pa
import urllib
import urllib.parse
import random
import smtplib
import time
status = True

def speak(text):
    speaker = win32com.client.Dispatch("SAPI.SpVoice")
    speaker.Speak(text)

def wishme():
    hour = int(datetime.datetime.now().hour)
    if hour >= 0 and hour < 12:
        speak("Good morning sir")
    elif hour >= 12 and hour < 18:
        speak("Good afternoon sir")
    else:
        speak("Good evening sir")

def record():
    chunk = 1024
    sample_format = pa.paInt16
    channels = 1
    fps = 44100
    filename = 'output.wav'
    global status
    p = pa.PyAudio()
    stream = p.open(format=sample_format, channels=channels, rate=fps, frames_per_buffer=chunk, input=True)
    frames = []
    while status:
        data = stream.read(chunk)
        frames.append(data)
    stream.stop_stream()
    stream.close()
    p.terminate()
    print("Successfully finished recording...")
    wf = wave.open(filename, 'wb')
    wf.setnchannels(channels)
    wf.setsampwidth(p.get_sample_size(sample_format))
    wf.setframerate(fps)
    wf.writeframes(b''.join(frames))
    wf.close()


def sendemail(to,content):
 server=smtplib.SMTP('smtp.gmail.com',587)
 server.ehlo()
 server.starttls()
 server.login('priyanshjyala2020@gmail.com','')
 server.sendmail('priyanshjyala2020@gmail.com',to,content)
 server.quit()

  
 




def takecommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        r.pause_threshold = 0.5
        speak("I am listening")
        speak("What you want me to do?")
        audio = r.listen(source)
        try:
            print("Recognizing...")
            query = r.recognize_google(audio, language="en-in")
            query = query.lower()
            print("User said: " + query)
            return query
        except Exception as e:
            print(e)
            return "some error occurred..."
        


if __name__ == "__main__":
    print("Engine starting...")
    speak("Initializing command mode")
    wishme()
    while True:
        query = takecommand()
        
        if query == "some error occurred...":
            continue

        sites = [
            ["youtube", "https://www.youtube.com"],
            ["stackoverflow","https://www.stackoverflow.com"],
            ["erp gehu","https://student.gehu.ac.in"],
            ["wikipedia", "https://www.wikipedia.com"],
            ["google", "https://www.google.com"],
            ["geeksforgeeks", "https://www.geeksforgeeks.org"],
            ["facebook", "https://www.facebook.com"],
            ["instagram", "https://www.instagram.com"]
        ]

        for site in sites:
            if f"open {site[0]}" in query:
                speak(f"Opening {site[0]} sir...")
                webbrowser.open(site[1])


        if "wikipedia"in query:
            speak("searching wikipedia...")
            query=query.replace("wikipedia","")
            result=wikipedia.summary(query)
            speak("according to wikipedia")
            print(result)
            speak(result)


        elif "web" in query:
            speak("searching webbrowser")
            query=urllib.parse.quote(query)
            url="https://www.google.com/?"+ query
            webbrowser.open_new(url)
            speak(webbrowser.open_new(url))
            opener=urllib.request.build_opener()
            response=opener.open(url)
            content=response.read().decode('utf-8')
            
            
        elif "play music" in query:
            dir= r"C:\Users\priya\Music"
            mp3music=[file for file in os.listdir(dir) if file.endswith('.mp3')]
            print(mp3music)
            random_music=random.choice(mp3music)
            music_file=os.startfile(os.path.join(dir,random_music))

        
            


        elif "what's the date and time" in query:
            hour = datetime.datetime.now().strftime("%H")
            mins = datetime.datetime.now().strftime("%M")
            date = datetime.datetime.now().strftime("%d")
            month= datetime.datetime.now().strftime("%m")
            year=  datetime.datetime.now().strftime("%Y")
            print(f"{date}-{month}-{year}")
            print(f"{hour}:{mins}")
            speak(f"date today is {date}-{month}-{year}")
            speak(f"and The time right now is {hour} hour {mins} minutes")




        elif "open visualstudiocode" in query:
            path=r"C:\Users\priya\AppData\Local\Programs\Microsoft VS Code\Code.exe"
            os.startfile(path)

        elif "open excel" in query:
            path=r"C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE"
            os.startfile(path)

        elif "open googlechrome" in query:
            path=r"C:\Program Files\Google\Chrome\Application\chrome.exe"
            os.startfile(path)


        elif "open word" in query:
            path=r"C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE"
            os.startfile(path)


elif "send email to me" in query:
    try:
        speak("what should i say")
        content=takecommand()
        to="priyanshsinghjyala813@gmail.com"
        sendemail(to,content)
        speak("email has been sent successfully")
    except Exception as e:
        print(e)
        speak("sorry,I m unable to send email")

        

elif "start recording" in query:
        thread1 = threading.Thread(target=record)
        thread1.start()

        if "stop recording" in query:
            status = False
            thread1.join()
            
            

    

