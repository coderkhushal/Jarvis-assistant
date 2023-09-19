import os

import win32com.client
import speech_recognition as sr
import webbrowser
import openai
import datetime
import requests
speaker= win32com.client.Dispatch("SAPI.SpVoice")
from AppOpener import open
import wikipedia

def ai(prompt):
    text = f"Open AI response for Prompt : {prompt}  \n**************************\n\n "
    try:

        # from apikeycontainer import apikey
        # openai.api_key = apikey

        response = openai.Completion.create(
            model="text-davinci-003",
            prompt=prompt,
            temperature=1,
            max_tokens=256,
            top_p=1,
            frequency_penalty=0,
            presence_penalty=0
        )

        print(response["choices"][0]["text"])
        if not os.path.exists("open_ai"):
            os.mkdir("open_ai")

        with open(f"open_ai/{' '.join(prompt.split(' ')[3:])}","a") as f:
            text += response["choices"][0]["text"]
            f.write(text)

        # {prompt.split('intelligence')[-1]}
        return "Done Sir"
    except Exception as e:
        print(e)
        return "Sorry Sir Some Error Occured"

chatstr = ""
def chat(query):
    global chatstr

    chatstr+= f"Khushal: {query}\n Jarvis: "
    try:

        # from apikeycontainer import apikey
        # openai.api_key = apikey

        response = openai.Completion.create(
            model="text-davinci-003",
            prompt=chatstr,
            temperature=1,
            max_tokens=256,
            top_p=1,
            frequency_penalty=0,
            presence_penalty=0
        )
        chatstr +=f" {response['choices'][0]['text']}\n "
        return response["choices"][0]["text"]
    except Exception as e:
        return "Sorry sir some error Occurred"

def takecommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        r.pause_threshold = 0.5
        audio = r.listen(source)

        try:
            query = r.recognize_google(audio, language="en-in")
            print(f"user said {query}")
            return query
        except  Exception  as e:
            return "some Error Occured Sir, I did not get what you said"
y="yes"
c=1


while "yes" in y or "ya" in y:
    if c==1:
        speaker.speak("JARVIS A i Sir. What can I do for you sir ")
    else:
        speaker.speak("OK sir waiting for your orders")
    query = takecommand()
    # speaker.speak(text)

    sites = [["youtube","https://www.youtube.com/results?search_query="," "], ["getbootstrap",], ["facebook"], ["amazon","https://www.amazon.in/s?k=","&crid=F97LT93XQP8W&sprefix=tshir%2Caps%2C467&ref=nb_sb_noss_2"], ["flipkart","https://www.flipkart.com/search?q=","&otracker=search&otracker1=search&marketplace=FLIPKART&as-show=on&as=off"],["instagram","","" ],["wikipedia"],["google","https://www.google.com/search?q=","&ei=JImnZKLpHNKw4-EPvt-s6AU&ved=0ahUKEwiii4Oj1vv_AhVS2DgGHb4vC10Q4dUDCA8&uact=5&oq=how+to+win+a+chess+game&gs_lp=Egxnd3Mtd2l6LXNlcnAiF2hvdyB0byB3aW4gYSBjaGVzcyBnYW1lMgUQABiABDIFEAAYgAQyBRAAGIAEMgUQABiABDIFEAAYgAQyBRAAGIAEMgUQABiABDIFEAAYgAQyBRAAGIAEMgUQABiABEiHL1C4B1jbLXADeAGQAQGYAZMCoAGkI6oBBjAuMjQuM7gBA8gBAPgBAagCFMICEBAAGIoFGOoCGLQCGEPYAQHCAhYQLhgDGI8BGOoCGLQCGIwDGOUC2AECwgIWEAAYAxiPARjqAhi0AhiMAxjlAtgBAsICBxAAGIoFGEPCAgsQABiABBixAxiDAcICCBAAGIoFGJECwgIFEC4YgATCAggQABiKBRixA8ICCBAuGLEDGIAEwgIXEC4YsQMYgAQYlwUY3AQY3gQY4ATYAQPCAgoQABgWGB4YDxgKwgIGEAAYFhge4gMEGAAgQYgGAboGBAgBGAe6BgYIAhABGAq6BgYIAxABGBQ&sclient=gws-wiz-serp"]]
    for site in sites:
        if f"open {site[0]}".lower() in query.lower() or f"search {site[0]}".lower() in query.lower():
            speaker.speak(f"what do you want to search in {site[0]} sir")
            srchwbsite = takecommand()
            speaker.speak(f"Opening {site[0]} sir...")
            if site[0]=="google":
                webbrowser.open(site[1]+ "+".join(srchwbsite.split(" ")) + site[2])
            else :
                webbrowser.open(site[1] + srchwbsite + site[2])
            break

    if"play" in query.lower() and "music" in query.lower() :
            speaker.speak("Playing Music Sir")
            musicPath = "C:\\Users\\hp\\Downloads\\G.O.A.T - Diljit Dosanjh.mp3"
            os.startfile(musicPath)
    elif "the time" in query.lower():
            hour = datetime.datetime.now().strftime("%H")
            minute = datetime.datetime.now().strftime("%M")

            if int(hour)<12:
                speaker.speak(f"Sir time is {hour} bajke {minute} minute ")
            else:
                speaker.speak(f"Sir time is {int(hour)-12} bajke {minute} minute ")
    elif "artificial intelligence" in query.lower():
        resp = ai(prompt= query[7:])
        speaker.speak(resp)

    elif "open" in query.lower() and "app" in query.lower():
        speaker.speak("Opening sir")
        # os.startfile(app[1])
        open(query.split("app")[-1], match_closest=True)
    elif "nothing" in query.lower():
        speaker.speak("Ok sir i think its better to keep my mouth shut ")
        y="no"
        break
    elif "boring" in query.lower():
        response = requests.get("https://www.boredapi.com/api/activity")

        speaker.speak("sir you can" + response.json()["activity"])
    elif"wikipedia" in query.lower():
        speaker.speak("searching wikipedia sir")
        try:
            result=wikipedia.summary(query.split("Wikipedia")[-1], sentences = 2 )
            speaker.speak(f"Sir according to wikipedia {result}")
        except Exception as e:
            speaker.speak("sorry sir there may be multiple references to this word. please provide detailed word")
            print(e)
    # else:
    #     speaker.speak("Sorry sir i did not got what you said ,  is this a website sir?")
    #     confirm = takecommand()
    #     if "yes" in confirm.lower():
    #         speaker.speak("Sir you have not added this website in the website list yet. should i add it")
    #         orders= takecommand()
    #         if 'yes' in orders.lower():
    #             sites.append(query.split("open")[-1])
    # else:
    #     resp = chat(query)
    #     speaker.speak(resp)

    speaker.speak("do you want any other help sir?")
    y= takecommand()
    if "yes" not in y and "ya" not in y:
        speaker.speak("OK sir, It was a pleasure helping you")
    c=2
