'''In String'''

# import requests

# City = input("Enter the name of the city: ")



# url = f"http://api.weatherapi.com/v1/current.json?key=2c895559a77e425485290026252708&q={City}"

# r =  requests.get(url)
# print(r.text)
# print(type(r.text))

'''In Dictionary'''

import requests
import json
import win32com.client

City = input("Enter the name of the city: ")

url = f"http://api.weatherapi.com/v1/current.json?key=2c895559a77e425485290026252708&q={City}"

speaker = win32com.client.Dispatch("SAPI.SpVoice")

r = requests.get(url)
dic = json.loads(r.text)
d = dic["current"]["temp_c"]
print(dic["location"]["country"])
print(dic["location"]["region"])
print(dic["location"]["name"])
print(dic["current"]["temp_c"])


print(f"The Current Weather in {City} is {d} degrees.")
speaker.Speak(f"The Current Weather in {City} is {d} degrees.")