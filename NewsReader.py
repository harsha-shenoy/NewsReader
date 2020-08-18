import requests
import json
import time
from win32com.client import Dispatch

def speak(str):
    speak = Dispatch("SAPI.SPVoice")
    speak.Speak(str)


if __name__ == '__main__':
    speak("News for today...Lets Begin...")
    url = "https://newsapi.org/v2/top-headlines?country=in&category=science&apiKey=743c18664f964f50befb91120ed07c9c"
    news = requests.get(url).text
    news_dict = json.loads(news)
    arts = news_dict['articles']
    for index,articles in enumerate(arts):
        if (index != len(arts)):
            speak("News Title")
            print(index,'\n', articles['title'])
            speak(articles['title'])
            speak("News description")
            print(index,'\n',articles['description'])
            speak(articles['description'])
            time.sleep(2)
            speak("Moving on to the next news")
        else:
            speak("That's all for today")
