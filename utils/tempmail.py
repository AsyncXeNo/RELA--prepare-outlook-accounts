import os
import re

import requests
from dotenv import load_dotenv

load_dotenv()

create_mailbox_url = 'https://flash-temp-mail.p.rapidapi.com/mailbox/create'
fetch_emails_url = 'https://flash-temp-mail.p.rapidapi.com/mailbox/emails'

X_RAPIDAPI_KEY = os.getenv('RAPIDAPI_KEY')
X_RAPIDAPI_HOST = os.getenv('RAPIDAPI_HOST')


def create_mailbox() -> dict:
    querystring = {'free_domains': 'false'}

    headers = {
        'x-rapidapi-key': X_RAPIDAPI_KEY,
        'x-rapidapi-host': X_RAPIDAPI_HOST,
        'content-type': 'application/json'
    }

    response = requests.post(create_mailbox_url, headers=headers, params=querystring)

    return response.json()


def fetch_emails(email: str) -> dict:
    querystring = { 'email_address': email }

    headers = {
        'x-rapidapi-key': X_RAPIDAPI_KEY,
        'x-rapidapi-host': X_RAPIDAPI_HOST,
        'content-type': 'application/json'
    }

    response = requests.get(fetch_emails_url, headers=headers, params=querystring)

    return response.json()


def extract_microsoft_otp(body: str) -> str:
    email_match = re.search(r'for the Microsoft account\s+([^\s.]+@[^\s.]+)', body)
    email = email_match.group(1) if email_match else None

    code_match = re.search(r'Security code:\s*(\d+)', body)
    code = code_match.group(1) if code_match else None

    if email and code:
        return f'{email} - {code}'
    else:
        return None
