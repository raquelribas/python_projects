import requests
import openpyxl
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
import pandas as pd
import datetime
import json

# API credentials
client_id = "your_id"
client_secret = "your_secret"
today = datetime.datetime.today()
today_param = today.strftime("%Y-%m-%d")
week_ago = today - datetime.timedelta(days=6)
week_ago_param = week_ago.strftime("%Y-%m-%d")
current_month = today.strftime("%B")
current_year = today.strftime("%Y")

def read_access_file():
    with open(r"C:\path\file.json","r") as json_file:
        jsonData = json.load(json_file)
    return jsonData

def overwrite_access_file(json_access):
    with open(r"C:\path\file.json","w") as json_file:
        json.dump(json_access, json_file)

def generate_token(id, secret, refresh):

    url = "https://api.rd.services/auth/token"

    payload = {
    "client_id": id,
    "client_secret": secret,
    "refresh_token": refresh
    }
    headers = {
        "accept": "application/json",
        "content-type": "application/json"
    }

    response = requests.post(url, json=payload, headers=headers)

    r = response.json()
    return r["access_token"], r["refresh_token"]

# info of all emails
def get_emails(access_token):

    headers = {
        "accept": "application/json",
        "authorization": f"Bearer {access_token}"
    }

    emails = []
    max_pages = 1
    page = 1
    url = f"https://api.rd.services/platform/emails?page={page}&page_size=125"

    response = requests.get(url, headers=headers)
    max_pages = int(response.headers['pagination-total-pages'])
    response = response.json()
    page +=1

    for item in response['items']:
        #item.pop('links')
        emails.append(item)
            
    
    while page <= max_pages:
        url = f"https://api.rd.services/platform/emails?page={page}&page_size=125"
        response = requests.get(url, headers=headers)
        response = response.json()

        page +=1
        
        for item in response['items']:
            #item.pop('links')
            emails.append(item)
    

    return emails

# analytics for email_mkt: open, clicks, etc.
def get_emailmkt_analytics(access_token):

    url = f"https://api.rd.services/platform/analytics/emails?start_date=2022-01-01&end_date={today_param}"

    headers = {
        "accept": "application/json",
        "authorization": f"Bearer {access_token}"
    }

    response = requests.get(url, headers=headers)
    emailmkt = response.json()

    return emailmkt['emails']

# analytics for workflow emails: open, clicks, etc.
def get_fluxo_analytics(access_token):

    url = f"https://api.rd.services/platform/analytics/workflow_emails?start_date={week_ago_param}&end_date={today_param}"

    headers = {
        "accept": "application/json",
        "authorization": f"Bearer {access_token}"
    }

    response = requests.get(url, headers=headers)
    fluxo = response.json()

    return fluxo['workflow_email_statistics']

# analytics for landing page coversion: visitors vs leads
def get_conversion_analytics(access_token):

    url = f"https://api.rd.services/platform/analytics/conversions?start_date={week_ago_param}&end_date={today_param}"

    headers = {
        "accept": "application/json",
        "authorization": f"Bearer {access_token}"
    }

    response = requests.get(url, headers=headers)
    conversion = response.json()

    return conversion['conversions']

# info of all landing pages
def get_landingpage(access_token):

    headers = {
        "accept": "application/json",
        "authorization": f"Bearer {access_token}"
    }

    landing_page = []
    max_pages = 1
    page = 1
    url = f"https://api.rd.services/platform/landing_pages?page={page}&page_size=125"

    response = requests.get(url, headers=headers)
    max_pages = int(response.headers['pagination-total-pages'])
    response = response.json()
    page +=1

    for item in response:
        landing_page.append(item)
            
    
    while page <= max_pages:
        url = f"https://api.rd.services/platform/landing_pages?page={page}&page_size=125"
        response = requests.get(url, headers=headers)
        response = response.json()

        page +=1
        
        for item in response:
            landing_page.append(item)

    return landing_page


read_json = read_access_file()
access_token = read_json["access_token"]
refresh_token = read_json["refresh_token"]


try:
    emails_response = get_emails(access_token)
except:
    access_token, refresh_token = generate_token(client_id, client_secret, refresh_token) 
    emails_response = get_emails(access_token)
    read_json["access_token"] = access_token
    read_json["refresh_token"] = refresh_token
    overwrite_access_file(read_json)

df = pd.DataFrame(emails_response)
df.to_excel(r"G:\path\Emails.xlsx", index=False)

emailmkt_analytics_response = get_emailmkt_analytics(access_token)

df = pd.DataFrame(emailmkt_analytics_response)
df.to_excel(r"G:\path\EmailMKT_analytics.xlsx", index=False)

fluxo_analytics_response = get_fluxo_analytics(access_token)

df = pd.DataFrame(fluxo_analytics_response)
df.to_excel(fr"G:\path\{current_year}\{current_month}\Fluxo_analytics_{today_param}.xlsx", index=False)

conversion_analytics_response = get_conversion_analytics(access_token)

df = pd.DataFrame(conversion_analytics_response)
df.to_excel(fr"G:\Drives compartilhados\Spinoff Drive\RD_Database\Landing_analytics\{current_year}\{current_month}\Conversion_analytics_{today_param}.xlsx", index=False)

landingpage_response = get_landingpage(access_token)

df = pd.DataFrame(landingpage_response)
df.to_excel(r"G:\path\landing_page.xlsx", index=False)
