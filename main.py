import speech_recognition as sr
import win32com.client

speaker = win32com.client.Dispatch('SAPI.SpVoice')

def speak(text):
    print(text)
    speaker.speak(text)

def takecommand():
    r = sr.Recognizer()
    with sr.Microphone(device_index = 0) as source:
        print('Listening...')
        r.pause_threshold = 1
        audio = r.listen(source)
        try:
            text = r.recognize_google(audio, language="en-in")
            print(f"user said:{text}")
            return text
        except Exception as e:
            print("Some Error occurred . Sorry")
    return ''

if __name__ == "__main__":
    while 1:
        command = takecommand()
        if "hello" in command:
            speak("Good Morning Boss How Can I Help You Today?")
        elif "your name" in command:
            speak("I am Jarvis Your Virtual Assistant")
        elif 'time' in command:
            from datetime import datetime
            now = datetime.now()
            current_time = now.strftime("%H:%M %p")
            speak(f"The Current Time is : {current_time}")
        elif 'date' in command:
            from datetime import datetime
            now = datetime.today()
            current_time = now.strftime("%Y-%m-%d")
            speak(f"The Date is : {current_time}")
        elif "exit" in command:
            speak("Good Night boss")
            break
        else:
            speak("I don't know how to response that, Please Try Again")


