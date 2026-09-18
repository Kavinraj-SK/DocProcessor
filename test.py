from google import genai
client = genai.Client(api_key="AQ.Ab8RN6IbMldLmsioDFlA30C5_uIwL2LoB4rIBUcyPp2O4_BBog")
for m in client.models.list():
    if "generateContent" in m.supported_actions:
        print(m.name)