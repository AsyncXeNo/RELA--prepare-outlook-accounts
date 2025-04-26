import os
import re
import email
import imaplib
from functools import lru_cache
from typing import List, Dict, Any, Optional
from email.utils import getaddresses
from email.policy import default
from collections import defaultdict
from email.policy import default

from loguru import logger

IMAP_HOST = os.getenv('IMAP_HOST')
IMAP_PORT = os.getenv('IMAP_PORT')

EMAIL_USER = os.getenv('EMAIL_USER')
EMAIL_PASS = os.getenv('EMAIL_PASS')

FEELD_LINK_PATTERN = re.compile(r'https://links\.fldcore\.com\?link=[^\s<>\"]+')
FOLDER_PATTERN = re.compile(r' "/" (.+?)$')

MAX_EMAILS_PER_BATCH = 50
MAX_BODY_SIZE = 50000
TRASH_FOLDER = 'Trash'


def connect() -> imaplib.IMAP4_SSL:
    try:
        mail = imaplib.IMAP4_SSL(IMAP_HOST, IMAP_PORT)
        mail.login(EMAIL_USER, EMAIL_PASS)
        return mail
    except Exception as e:
        logger.error(f'Connection failed: {e}')
        raise


@lru_cache(maxsize=1)
def get_folder_names(mail: imaplib.IMAP4_SSL) -> List[str]:
    try:
        _, folders = mail.list()
        folder_names = []
        for folder in folders:
            folder_str = folder.decode()
            match = FOLDER_PATTERN.search(folder_str)
            if match:
                folder_name = match.group(1).strip('"')
                folder_names.append(folder_name)
        return folder_names
    except Exception as e:
        logger.error(f'Error getting folder names: {e}')
        return []
    

def extract_email_parts(msg: email.message.Message) -> Dict[str, Any]:
    from_header = msg.get('From', '')
    from_email = getaddresses([from_header])
    from_email = from_email[0][1] if from_email else ''
    
    to_addresses = [addr for _, addr in getaddresses([msg.get('To', '')])]
    
    subject = msg.get('Subject', '')
    date_header = msg.get('Date', '')
    
    return {
        'from': from_email,
        'to': to_addresses,
        'subject': subject,
        'date': date_header
    }


def extract_text_payload(part: email.message.Message) -> str:
    if part.get_content_type() == 'text/plain' and 'attachment' not in str(part.get('Content-Disposition', '')):
        payload = part.get_payload(decode=True)
        if payload:
            return payload[:MAX_BODY_SIZE].decode(errors='ignore')
    return ''


def get_email_body(msg: email.message.Message) -> str:
    body_parts = []
    
    if msg.is_multipart():
        for part in msg.walk():
            body_part = extract_text_payload(part)
            if body_part:
                body_parts.append(body_part)
    else:
        payload = msg.get_payload(decode=True)
        if payload:
            body_parts.append(payload[:MAX_BODY_SIZE].decode(errors='ignore'))
            
    return ''.join(body_parts)


def extract_microsoft_otp(body: str) -> str:
    email_match = re.search(r'for the Microsoft account\s+([^\s.]+@[^\s.]+)', body)
    email = email_match.group(1) if email_match else None

    code_match = re.search(r'Security code:\s*(\d+)', body)
    code = code_match.group(1) if code_match else None

    if email and code:
        return f'{email} - {code}'
    else:
        return None


def fetch_emails_batch(mail: imaplib.IMAP4_SSL, uids: List[bytes]) -> List[Dict[str, Any]]:
    emails = []
    if not uids:
        return emails
        
    uid_str = b','.join(uids).decode()
    
    try:
        result, data = mail.uid('fetch', uid_str, '(BODY.PEEK[])')
        if result != 'OK':
            return emails
            
        i = 0
        while i < len(data):
            if isinstance(data[i], tuple):
                header = data[i][0].decode()
                uid_match = re.search(r'UID (\d+)', header)
                if not uid_match:
                    i += 1
                    continue
                    
                uid = uid_match.group(1)
                raw_email = data[i][1]
                
                msg = email.message_from_bytes(raw_email, policy=default)
                
                email_parts = extract_email_parts(msg)
                body = get_email_body(msg)
                microsoft_otp = extract_microsoft_otp(body)
                
                emails.append({
                    'uid': uid,
                    'from': email_parts['from'],
                    'to': email_parts['to'],
                    'subject': email_parts['subject'],
                    'body': body,
                    'microsoft_otp': microsoft_otp,
                    'date': email_parts['date'],
                })
            i += 1
    except Exception as e:
        logger.error(f'Error fetching email batch: {e}')
        
    return emails
    

def get_emails_by_folder(mail: imaplib.IMAP4_SSL, folder_name: str, 
                         max_emails: int = 999, 
                         search_criteria: str = 'ALL') -> List[Dict[str, Any]]:
    all_emails = []
    
    try:
        mail.select(f'"{folder_name}"', readonly=True)
        
        result, data = mail.uid('search', None, search_criteria)
        if result != 'OK' or not data[0]:
            return []

        uids = data[0].split()
        
        uids = uids[:max_emails]
        
        for i in range(0, len(uids), MAX_EMAILS_PER_BATCH):
            batch_uids = uids[i:i+MAX_EMAILS_PER_BATCH]
            batch_emails = fetch_emails_batch(mail, batch_uids)
            
            for email_data in batch_emails:
                email_data['folder'] = folder_name
                
            all_emails.extend(batch_emails)
            
    except Exception as e:
        logger.error(f'Error processing folder {folder_name}: {e}')

    return all_emails


def get_all_emails(mail: imaplib.IMAP4_SSL, 
                   max_emails_per_folder: int = 999,
                   folders: Optional[List[str]] = None,
                   search_criteria: str = 'ALL') -> List[Dict[str, Any]]:
    all_emails = []
    
    if folders is None:
        folders = get_folder_names(mail)
        
    folders = [folder for folder in folders if folder != TRASH_FOLDER]
    
    for folder_name in folders:
        folder_emails = get_emails_by_folder(
            mail, 
            folder_name, 
            max_emails=max_emails_per_folder,
            search_criteria=search_criteria
        )
        all_emails.extend(folder_emails)

    return all_emails


def move_emails_to_trash(mail: imaplib.IMAP4_SSL, emails: List[Dict[str, Any]]) -> None:
    if not emails:
        return
    
    folder_emails = defaultdict(list)
    for email in emails:
        folder_emails[email['folder']].append(email['uid'])
    
    for folder_name, uids in folder_emails.items():
        try:
            mail.select(f'"{folder_name}"', readonly=False)
            
            uid_str = ','.join(uids)
            
            result, _ = mail.uid('COPY', uid_str, f'"{TRASH_FOLDER}"')
            if result != 'OK':
                logger.error(f'Failed to copy emails to Trash from {folder_name}')
                continue
                
            mail.uid('STORE', uid_str, '+FLAGS', '\\Deleted')
            
            mail.expunge()
            
        except Exception as e:
            logger.error(f'Error moving emails to Trash from {folder_name}: {e}')


def move_email_to_trash(mail: imaplib.IMAP4_SSL, email: Dict[str, Any]) -> None:
    move_emails_to_trash(mail, [email])