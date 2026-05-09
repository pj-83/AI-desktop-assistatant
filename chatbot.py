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
import random
import smtplib
import time
import sounddevice as sd
import numpy as np
import scipy.io.wavfile as wav

status = True

# ---------------- SPEAK ----------------
def speak(text):
    speaker = win32com.client.Dispatch("SAPI.SpVoice")
    speaker.Speak(text)

# ---------------- WISH ----------------
def wishme():
    hour = int(datetime.datetime.now().hour)
    if hour < 12:
        speak("Good morning sir")
    elif hour < 18:
        speak("Good afternoon sir")
    else:
        speak("Good evening sir")

# ---------------- RECORD ----------------
def record():
    global status
    fs = 44100
    filename = "output.wav"
    recording = []

    def callback(indata, frames, time, status_):
        if status:
            recording.append(indata.copy())

    with sd.InputStream(samplerate=fs, channels=1, callback=callback):
        print("Recording started...")
        while status:
            sd.sleep(100)

    audio = np.concatenate(recording, axis=0)
    audio = (audio * 32767).astype(np.int16)

    wf = wave.open(filename, 'wb')
    wf.setnchannels(1)
    wf.setsampwidth(2)
    wf.setframerate(fs)
    wf.writeframes(audio.tobytes())
    wf.close()

    print("Recording saved")

# ---------------- TAKE COMMAND ----------------
def takecommand():
    r = sr.Recognizer()
    r.energy_threshold = 300
    r.dynamic_energy_threshold = True

    fs = 16000
    duration = 4

    speak("I am listening")
    print("Listening...")

    recording = sd.rec(int(duration * fs), samplerate=fs, channels=1)
    sd.wait()

    # 🔥 FIX: convert float → int16 (VERY IMPORTANT)
    recording = (recording * 32767).astype(np.int16)

    # save proper PCM WAV
    wav.write("temp.wav", fs, recording)

    try:
        with sr.AudioFile("temp.wav") as source:
            audio = r.record(source)

        print("Recognizing...")
        query = r.recognize_google(audio, language="en-in")
        query = query.lower()

        print("User said:", query)
        return query

    except Exception as e:
        print(e)
        speak("Sorry, I did not understand")
        return "some error occurred..."

# ---------------- EMAIL ----------------
def sendemail(to, content):
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.starttls()
    server.login('your_email@gmail.com', 'your_app_password')
    server.sendmail('your_email@gmail.com', to, content)
    server.quit()

# ---------------- MAIN ----------------
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
            ["stackoverflow", "https://www.stackoverflow.com"],
            ["erp gehu", "https://student.gehu.ac.in"],
            ["wikipedia", "https://www.wikipedia.com"],
            ["google", "https://www.google.com"],
            ["geeksforgeeks", "https://www.geeksforgeeks.org"],
            ["facebook", "https://www.facebook.com"],
            ["instagram", "https://www.instagram.com"]
        ]

        for site in sites:
            if f"open {site[0]}" in query:
                speak(f"Opening {site[0]}")
                webbrowser.open(site[1])

        if "wikipedia" in query:
            speak("searching wikipedia")
            query = query.replace("wikipedia", "")
            result = wikipedia.summary(query, sentences=2)
            speak(result)
            print(result)

        elif "web" in query:
            speak("searching webbrowser")
            query = urllib.parse.quote(query)
            url = "https://www.google.com/search?q=" + query
            webbrowser.open_new(url)

        elif "play music" in query:
            dir = r"C:\Users\priya\Music"
            mp3music = [file for file in os.listdir(dir) if file.endswith('.mp3')]
            if mp3music:
                random_music = random.choice(mp3music)
                os.startfile(os.path.join(dir, random_music))

        elif "what's the date and time" in query:
            now = datetime.datetime.now()
            speak(f"date today is {now.strftime('%d-%m-%Y')}")
            speak(f"time is {now.strftime('%H')} hour {now.strftime('%M')} minutes")

        elif "open visualstudiocode" in query:
            os.startfile(r"C:\Users\priya\AppData\Local\Programs\Microsoft VS Code\Code.exe")

        elif "open excel" in query:
            os.startfile(r"C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE")

        elif "open googlechrome" in query:
            os.startfile(r"C:\Program Files\Google\Chrome\Application\chrome.exe")

        elif "open word" in query:
            os.startfile(r"C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE")

        elif "send email to me" in query:
            try:
                speak("what should i say")
                content = takecommand()
                sendemail("your_email@gmail.com", content)
                speak("email sent successfully")
            except Exception as e:
                print(e)
                speak("unable to send email")

        elif "start recording" in query:
            status = True
            threading.Thread(target=record).start()

        elif "stop recording" in query:
            status = False

        elif "exit" in query:
            speak("Goodbye")
            break
