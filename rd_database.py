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
today = datetime.datetime.today().strftime("%Y-%m-%d %H:%M:%S")
today2 = datetime.datetime.today().strftime("%Y-%m-%d")

def read_access_file():
    with open(r"C:\Users\admin\Desktop\Python\Codes\rd_tokens.json","r") as json_file:
        jsonData = json.load(json_file)
    return jsonData

def overwrite_access_file(json_access):
    with open(r"C:\Users\admin\Desktop\Python\Codes\rd_tokens.json","w") as json_file:
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

# list of all emails
def get_emails(access_token):

    print(f"Executando get_emails {today}")
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
    print(f"get_emails concluída! {today}")

    return emails

# Lista de todos os uuid na segmentação 'Todos os contatos da Base de Leads'
df = pd.read_excel(r"G:\path\file.xlsx", sheet_name=0)
list_uuid = df['uuid'].tolist()


# Detalhes de todos os contatos da base - não executar sempre!
def get_contact_details(access_token, list_uuid):

    print(f"Executando get_contact_details {today}")    
    contact_details = []

    for uuid in list_uuid:

        url = f"https://api.rd.services/platform/contacts/uuid:{uuid}"

        headers = {
            "accept": "application/json",
            "authorization": f"Bearer {access_token}"
        }

        response = requests.get(url, headers=headers)
        contact_details.append(response.json())
    print(f"get_contact_details concluída! {today}")

    return contact_details

def get_funil_contato(access_token, list_uuid):

    print(f"Executando get_funil_contato {today}")
    contact_funil = []

    for uuid in list_uuid:

        url = f"https://api.rd.services/platform/contacts/uuid:{uuid}/funnels/default"

        headers = {
            "accept": "application/json",
            "authorization": f"Bearer {access_token}"
        }

        response = requests.get(url, headers=headers)
        contact_funil.append(response.json())
    print(f"get_funil_contato concluída! {today}")

    return contact_funil

def get_contact_events(access_token, list_uuid):

    print(f"Executando get_contact_events {today}")    
    contact_events = []

    for uuid in list_uuid:
        url = f"https://api.rd.services/platform/contacts/{uuid}/events"

        headers = {
            "accept": "application/json",
            "authorization": f"Bearer {access_token}"
        }

        params = {
            "event_type": ["CONVERSION", "OPPORTUNITY"]
        }

        response = requests.get(url, headers=headers)
        contact_events.append(response.json())
    print(f"get_contact_events concluída! {today}")

    return response.json()

def get_fluxoauto(access_token):

    print(f"Executando get_fluxoauto {today}") 
    headers = {
        "accept": "application/json",
        "authorization": f"Bearer {access_token}"
    }

    fluxo_auto = []
    max_pages = 1
    page = 1
    url = f"https://api.rd.services/platform/workflows?page={page}&page_size=125"

    response = requests.get(url, headers=headers)
    max_pages = int(response.headers['pagination-total-pages'])
    response = response.json()
    page +=1

    for item in response['workflows']:
        #item.pop('links')
        fluxo_auto.append(item)
            
    
    while page <= max_pages:
        url = f"https://api.rd.services/platform/workflows?page={page}&page_size=125"
        response = requests.get(url, headers=headers)
        response = response.json()

        page +=1
        
        for item in response['workflows']:
            #item.pop('links')
            fluxo_auto.append(item)
    print(f"get_fluxoauto concluída! {today}")

    return fluxo_auto


def get_fluxoauto_details(access_token, list_fluxoid):

    print(f"Executando get_fluxoauto_details {today}")
    fluxoauto_details = []

    for id in list_fluxoid:

        url = f"https://api.rd.services/platform/workflows/{id}"

        headers = {
            "accept": "application/json",
            "authorization": f"Bearer {access_token}"
        }

        response = requests.get(url, headers=headers)
        fluxoauto_details.append(response.json())
    print(f"get_fluxoauto_details concluída! {today}")

    return fluxoauto_details

def get_emailmkt_analytics(access_token):

    url = f"https://api.rd.services/platform/analytics/emails?start_date=2022-01-01&end_date={today2}"

    headers = {
        "accept": "application/json",
        "authorization": f"Bearer {access_token}"
    }

    response = requests.get(url, headers=headers)
    emailmkt = response.json()

    return emailmkt['emails']

def get_fluxo_analytics(access_token):

    url = f"https://api.rd.services/platform/analytics/workflow_emails?start_date=2022-01-01&end_date={today2}"

    headers = {
        "accept": "application/json",
        "authorization": f"Bearer {access_token}"
    }

    response = requests.get(url, headers=headers)
    fluxo = response.json()

    return fluxo['workflow_email_statistics']

def get_funil_analytics(access_token):

    url = f"https://api.rd.services/platform/analytics/funnel?start_date=2023-01-01&end_date={today2}&grouped_by=weekly"

    headers = {
        "accept": "application/json",
        "authorization": f"Bearer {access_token}"
    }

    response = requests.get(url, headers=headers)
    funil = response.json()

    return funil['funnel']

def get_conversion_analytics(access_token):

    url = f"https://api.rd.services/platform/analytics/conversions?start_date=2023-01-01&end_date={today2}"

    headers = {
        "accept": "application/json",
        "authorization": f"Bearer {access_token}"
    }

    response = requests.get(url, headers=headers)
    conversion = response.json()

    return conversion['conversions']

def get_landingpage(access_token):

    print(f"Executando get_landingpage {today}") 
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

    print(f"get_landingpage concluída! {today}")

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
df.to_excel(r"G:\Drives compartilhados\Spinoff Drive\RD_Database\Emails.xlsx", index=False)

'''fluxo_auto_response = get_fluxoauto(access_token)

df = pd.DataFrame(fluxo_auto_response)
df.to_excel(r"G:\Drives compartilhados\Spinoff Drive\RD_Database\FluxoAUTO.xlsx", index=False)

df = pd.read_excel(r"G:\Drives compartilhados\Spinoff Drive\RD_Database\FluxoAUTO.xlsx", sheet_name=0)
list_fluxoid = df['id'].tolist()

fluxoauto_details_response = get_fluxoauto_details(access_token, list_fluxoid)

df = pd.DataFrame(fluxoauto_details_response)
df.to_excel(r"G:\Drives compartilhados\Spinoff Drive\RD_Database\FluxoAUTO_details.xlsx", index=False)'''

emailmkt_analytics_response = get_emailmkt_analytics(access_token)

df = pd.DataFrame(emailmkt_analytics_response)
df.to_excel(r"G:\Drives compartilhados\Spinoff Drive\RD_Database\EmailMKT_analytics.xlsx", index=False)

funil_analytics_response = get_funil_analytics(access_token)

df = pd.DataFrame(funil_analytics_response)
df.to_excel(r"G:\Drives compartilhados\Spinoff Drive\RD_Database\Funil_analytics.xlsx", index=False)

landingpage_response = get_landingpage(access_token)

df = pd.DataFrame(landingpage_response)
df.to_excel(r"G:\Drives compartilhados\Spinoff Drive\RD_Database\landing_page.xlsx", index=False)
