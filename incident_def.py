import pandas as pd
from pyshelle import ShelleClient
import time
import requests
from requests.exceptions import ReadTimeout

APPLICATION_ID = 47
CLIENT_ID = "a35c0449-2edc-44aa-ac89-2b15c1406aa8"
CLIENT_PASS = "3669d7cb-0379-4496-9ad0-45e4af2cb586"
CLIENT_SECRET = "03a128e1-6bad-4841-8a12-6f9ebf839d1b"
NEW_ENDPOINT = 'https://nprd-sbtst-shelleapimgmt.azure-api.net/backend'
SUBSCRIPTION_KEY = 'ea08e1381ab94c779e5c4937ccb6162e'

client = ShelleClient(APPLICATION_ID, CLIENT_ID, CLIENT_PASS, CLIENT_SECRET,
    endpoint=NEW_ENDPOINT,
    subscription_key=SUBSCRIPTION_KEY,
    proxies={
      'http': 'zproxy-global.shell.com:80',
      'https': 'zproxy-global.shell.com:80'
    }
)

system_prompt = """Can you act as a sentiment score analyser, on a scale of 1-10, with 10 being the most unhappy. I will give you incident information, you have to assign a sentiment score to it. Mainly the Description and Comments and Work notes would be an indicator of the user's sentiment, please give accurate indication of their satisfaction or frustration. You are ONLY supposed to reply in digits, that is between 1-10. Alphabets are not accepted in your response."""
overrides = {"prompt" : system_prompt,"temperature": 0.0}

def get_sentiment_score(text, retries=3, backoff_factor=0.3):
    prompt = f"{system_prompt} {text}"
    for attempt in range(retries):
        try:
            response = client.get_response(prompt, overrides=overrides, timeout=90)
            if not response.error:
                return int(response.message)
            else:
                return None
        except ReadTimeout:
            if attempt < retries - 1:
                time.sleep(backoff_factor * (2 ** attempt))
                continue
            else:
                print(f"Request timed out after {retries} attempts.")
                return None
            
df = pd.read_excel('altered_incident.xlsx', engine='openpyxl')

df['Combined_Text'] = df['Description'] + ' ' + df['Comments and Work notes']

sentiment_scores = []
for text in df['Combined_Text']:
    sentiment_score = get_sentiment_score(text)
    sentiment_scores.append(sentiment_score)

df['Sentiment_Score'] = sentiment_scores

output_file_path = "C:\\Users\\Avni.Tonger\\Downloads\\sentimentanalysis4.xlsx"
df.to_excel(output_file_path, index=False)

print("Incidents are saved in the new excel file")