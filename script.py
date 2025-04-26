import config as _

import os
import time

from loguru import logger
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from utils.airtable import get_entries, update_entry
from utils.selenium import create_driver, quit_driver
from utils.email import connect, get_all_emails, move_email_to_trash

MICROSOFT_LOGIN_URL = 'https://login.live.com/login.srf'
JUNK_EMAIL = 'https://outlook.live.com/mail/0/options/mail/junkEmail'
FORWARDING = 'https://outlook.live.com/mail/0/options/mail/forwarding'
FEELD_EMAIL = 'noreply@open.feeld.co'

EMAIL_USER = os.getenv('EMAIL_USER')


def main():

    logger.info('Starting script.')
    
    logger.debug('Fetching entries from Airtable.')
    entries = get_entries()

    logger.debug(f'Fetched {len(entries)} entries from Airtable.')

    for index, entry in enumerate(entries[:3]):

        logger.info(f'[{index+1}/{len(entries)}] Processing email {entry.get("Email")}')

        if entry.get('Prepared'):
            logger.warning(f'[{index+1}/{len(entries)}] {entry.get("Email")} has already been prepared. Skipping.')
            continue

        driver = create_driver()

        # Logging in and adding recovery email

        logger.debug(f'[{index+1}/{len(entries)}] Logging in')
        
        driver.get(MICROSOFT_LOGIN_URL)

        try:
            WebDriverWait(driver, 3).until(
                EC.visibility_of_element_located(
                    (By.CSS_SELECTOR, '#usernameEntry')
                )
            ).send_keys(entry.get('Email'))
        except Exception:
            WebDriverWait(driver, 3).until(
                EC.visibility_of_element_located(
                    (By.CSS_SELECTOR, '#i0116')
                )
            ).send_keys(entry.get('Email'))
        time.sleep(0.2)

        WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable(
                (By.CSS_SELECTOR, 'button[type="submit"]')
            )
        ).click()

        try:
            WebDriverWait(driver, 3).until(
                EC.visibility_of_element_located(
                    (By.CSS_SELECTOR, '#passwordEntry')
                )
            ).send_keys(entry.get('Password'))
        except Exception:
            WebDriverWait(driver, 3).until(
                EC.visibility_of_element_located(
                    (By.CSS_SELECTOR, '#i0118')
                )
            ).send_keys(entry.get('Password'))
        time.sleep(0.2)

        url = driver.current_url

        WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable(
                (By.CSS_SELECTOR, 'button[type="submit"]')
            )
        ).click()

        while url == driver.current_url:
            time.sleep(0.2)

        if 'privacynotice' in driver.current_url:
            WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable(
                    (By.CSS_SELECTOR, 'button[type="button"]')
                )
            ).click()

            logger.debug(f'[{index+1}/{len(entries)}] Privacy notice accepted.')

        WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable(
                (By.XPATH, '//button[normalize-space()="Yes"]')
            )
        ).click()

        if not entry.get('Recovery Email'):

            logger.debug(f'[{index+1}/{len(entries)}] Waiting for recovery email input.')

            WebDriverWait(driver, 15).until(
                EC.visibility_of_element_located(
                    (By.CSS_SELECTOR, '#EmailAddress')
                )
            ).send_keys(EMAIL_USER)
            time.sleep(0.2)
            
            WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable(
                    (By.CSS_SELECTOR, '#iNext')
                )
            ).click()

            logger.debug(f'[{index+1}/{len(entries)}] Fetching OTP.')

            while True:
                try:
                    mail = connect()
                    break
                except Exception as e:
                    time.sleep(5)

            verification_code = None

            while True:
                emails = get_all_emails(mail)
                if not emails:
                    time.sleep(3)
                    continue
                for email in emails:
                    if email.get('microsoft_otp'):
                        otp = email.get('microsoft_otp')
                        print(otp)
                        email_address = otp.split('-')[0].strip().split('@')[0].strip()
                        code = otp.split('-')[-1].strip()
                        print(code)

                        email_start = email_address.split('*')[0]
                        email_end = email_address.split('*')[-1]

                        if entry.get('Email').split('@')[0].strip().startswith(email_start) and entry.get('Email').split('@')[0].strip().endswith(email_end):
                            verification_code = code
                            move_email_to_trash(mail, email)
                            break

                if verification_code:
                    break

                time.sleep(3)

            logger.debug(f'[{index+1}/{len(entries)}] OTP: {verification_code}')

            WebDriverWait(driver, 5).until(
                EC.visibility_of_element_located(
                    (By.CSS_SELECTOR, '#iOttText')
                )
            ).send_keys(verification_code)
            time.sleep(0.2)

            WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable(
                    (By.CSS_SELECTOR, '#iNext')
                )
            ).click()
            time.sleep(0.2)

            update_entry(entry.get('id'), {'Recovery Email': EMAIL_USER})

            logger.debug(f'[{index+1}/{len(entries)}] Recovery email added.')

        logger.debug(f'[{index+1}/{len(entries)}] Logged in successfully.')
        
        # Adding safe sender

        driver.get(JUNK_EMAIL)

        try:
            WebDriverWait(driver, 20).until(
                EC.element_to_be_clickable(
                    (By.XPATH, '//button[normalize-space()="No, thanks"]')
                )
            ).click()
        except Exception:
            pass

        if not entry.get('Safe Sender Added'):
            logger.debug(f'[{index+1}/{len(entries)}] Adding safe sender.')
            driver.get(JUNK_EMAIL)

            WebDriverWait(driver, 20).until(
                EC.element_to_be_clickable(
                    (By.XPATH, '//button[normalize-space()="Add safe sender"]')
                )
            ).click()
            time.sleep(0.2)
            
            WebDriverWait(driver, 5).until(
                EC.visibility_of_all_elements_located(
                    (By.CSS_SELECTOR, '.fui-Input__input')
                )
            )[-1].send_keys(FEELD_EMAIL)
            time.sleep(0.2)
            
            WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable(
                    (By.XPATH, '//button[normalize-space()="OK"]')
                )
            ).click()
            time.sleep(0.2)

            WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable(
                    (By.XPATH, '//button[normalize-space()="Save"]')
                )
            ).click()
            time.sleep(0.2)

            update_entry(entry.get('id'), {'Safe Sender Added': True})

            logger.debug(f'[{index+1}/{len(entries)}] Safe sender added.')

        # Adding forwarding email

        if not entry.get('Forwarded Email'):

            logger.debug(f'[{index+1}/{len(entries)}] Adding forwarding email.')

            driver.get(FORWARDING)

            WebDriverWait(driver, 30).until(
                EC.element_to_be_clickable(
                    (By.CSS_SELECTOR, '.fui-Switch')
                )
            ).click()
            time.sleep(0.2)

            WebDriverWait(driver, 5).until(
                EC.visibility_of_all_elements_located(
                    (By.CSS_SELECTOR, '.fui-Input__input')
                )
            )[-1].send_keys(EMAIL_USER)
            time.sleep(0.2)

            WebDriverWait(driver, 5).until(
                EC.presence_of_all_elements_located(
                    (By.CSS_SELECTOR, '.ms-Checkbox-checkbox')
                )
            )[-1].click()
            time.sleep(0.2)

            WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable(
                    (By.XPATH, '//button[normalize-space()="Save"]')
                )
            ).click()
            time.sleep(0.2)

            update_entry(entry.get('id'), {'Forwarded Email': EMAIL_USER})

            logger.debug(f'[{index+1}/{len(entries)}] Forwarding email added.')

        update_entry(entry.get('id'), {'Prepared': True})

        quit_driver()

        logger.info('Waiting for 10 seconds before next entry.')

        time.sleep(10)


if __name__ == '__main__':
    main()