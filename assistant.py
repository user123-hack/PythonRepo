import os
import webbrowser
import speech_recognition as s_r
import pyttsx3
import espeakng

class Assistant:
    def __init__(self):
        
        question = "\N{thinking face}" * 28
        face = "\N{thinking face}" * 5  

        engine = pyttsx3.init()
        engine.setProperty('rate', 110)      
        engine.say("Welcome Master, What Can I Do For You Today?")
        engine.runAndWait()

        print(question + "\n" + question)
        print( face + " Welcome User, What Can I Do For You " + face + "\n" + question + "\n" + question)

        engine = pyttsx3.init()
        engine.setProperty('rate', 110)      
        engine.say("Enter Your Preference ")
        engine.runAndWait()

        option = int(input("\nEnter Your Choice \n 1. Voice Based " + "\N{microphone}" + "\n 2. Text Based " + "\N{keyboard}" + "\n\n"))

        self.choice(option)

    def speech(self, item):
            engine = pyttsx3.init()
            engine.setProperty('rate', 110)      
            
            if(item == ""):
                engine.say("Enter A Valid Choice")
            
            else:
                engine.say("Opening " + item)

            engine.runAndWait()
    
    def search(self, input):
        input = input.lower()

        if (input == ""):
            print("enter valid choice" + "\N{thinking face}")
            self.speech("")

        elif("brave" in input):
            print("\n Opening Brave Browser" + " \N{Globe with Meridians}")
            self.speech("brave browser")
            os.system("brave-browser")
        
        elif("android" in input):
            print("\n Opening Studio " + "\N{fire}" *5)
            self.speech("Android Studio")
            os.system("android-studio")

        elif("telegram" in input):
            print("\n Opening Telegram " + "\N{Envelope}")
            self.speech("Telegram Desktop")
            os.system("telegram-desktop")

        elif("office" in input or "libre" in input):
            self.speech("Libre Office")
            os.system("libreoffice")

        elif("search" in input):
            key = input[6::]
            print("\n Searching" + key + "on"  + " \N{Globe with Meridians}")
            self.speech("Searching" + key + "on your browser")
            webbrowser.open_new_tab('http://www.google.com/search?btnG=1&q=%s' %key)

        elif("anydesk" in input or "desk" in input):
            self.speech("AnyDesk")
            os.system("libreoffice")

        elif("filezilla" in input or "zilla" in input):
            self.speech("FileZilla")
            os.system("filezilla")    

        elif("camera" in input):
            self.speech("Camera")
            os.system("cheese")

        elif("teamviewer" in input):
            self.speech("TeamViewer")
            os.system("teamviewer")

        elif("code" in input or "editor" in input or "visual" in input or "studio" in input):
            self.speech("Visual Studio Code")
            os.system("code")    

        elif("vlc" in input):
            self.speech("VLC Media Player")        
            os.system("vlc")

        elif("teams" in input):
            self.speech("Mircrosoft Teams")        
            os.system("teams")
        
    def choice(self,a):
        if(a == 1):
            
            mic = s_r.Microphone()
            r = s_r.Recognizer()
            my_mic = s_r.Microphone()

            with my_mic as source:
                print("Speak Your Desire")
                r.pause_threshold = 1
                audio = r.listen(source)
                r.adjust_for_ambient_noise(source)
                
                try:
                    print("Recognizing...")    
                    query = r.recognize_google(audio, language='en-in') 
                    print(f"User said: {query}\n")
                    self.search(query)
                    
                except Exception as e:
                    print("Please Say That Again")

        
        elif(a == 2):
            
            data = input("\nWhat Can I Do For You ?????\n")
            self.search(data)    

if __name__ == "__main__":
    a = Assistant()            