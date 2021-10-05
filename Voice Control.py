import win32com
import pyttsx3
import speech_recognition as mr
import webbrowser
import os
import os.path
import win32com.client as wincl


engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[1].id)

def speak(audio):
    engine.say(audio)
    engine.runAndWait()

def wishMe():
    speak("Welcome Sir!!!. How can i help you today")

def takeCommand():
    r = mr.Recognizer()
    with mr.Microphone() as source:
        print("Listening...")
        r.pause_threshold = 1
        audio = r.listen(source)

    try:
        print("Recognizing...")
        query = r.recognize_google(audio, language='en-in')
        print(f"Input Received: {query}\n")

    except Exception as e:
        print("Waiting for your comment...")
        speak("Waiting for your comment")
        return "None"
    return query

if __name__ == "__main__":
    wishMe()
    while True:

        query = takeCommand().lower()
        if 'open linkedin' in query:
            webbrowser.open("https://www.linkedin.com/signup")
            speak("Web Application is ready to use now...please tell me if i can assist further")

        elif 'emails' in query:
            try:
                if os.path.exists('E:\\desktop\\Test_VBA.xlsm'):
                    excel_macro = win32com.client.DispatchEx(
                        "Excel.Application") 
                    excel_path = os.path.expanduser('E:\\desktop\\Test_VBA.xlsm')
                    workbook = excel_macro.Workbooks.Open(Filename=excel_path, ReadOnly=1)
                    excel_macro.Application.Run(
                        "Test_VBA.xlsm!Button1_Click")  
                    workbook.Save()
                    excel_macro.Application.Quit()
                    del excel_macro
                    speak("Execution completed, How can i help you further")
            except Exception as f:     
                print(f)
                speak("Execution not completed, please tell me how can i help you further")

        elif 'open code' in query:
            codePath = "E:\\desktop\\Test_VBA.xlsm"
            os.startfile(codePath)
            speak("File is ready to use now...please tell me if i can assist further")




