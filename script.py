import config as _

import os
import time

import traceback
import pyperclip
from loguru import logger
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from utils.airtable import get_entries, update_entry
from utils.selenium import create_driver, quit_driver
from utils.tempmail import create_mailbox, fetch_emails, extract_microsoft_otp

MICROSOFT_LOGIN_URL = 'https://login.live.com/login.srf'
JUNK_EMAIL = 'https://outlook.live.com/mail/0/options/mail/junkEmail'
FORWARDING = 'https://outlook.live.com/mail/0/options/mail/rules'
FEELD_EMAIL = 'noreply@open.feeld.co'
TAIMI_EMAIL = 'activate@taimi.com'

EMAIL_USER = os.getenv('EMAIL_USER')


def main():

    logger.info('Starting script.')
    
    logger.debug('Fetching entries from Airtable.')
    entries = get_entries()

    logger.debug(f'Fetched {len(entries)} entries from Airtable.')

    for index, entry in enumerate(entries):

        logger.info(f'[{index+1}/{len(entries)}] Processing email {entry.get("Email")}')

        if entry.get('Prepared'):
            logger.warning(f'[{index+1}/{len(entries)}] {entry.get("Email")} has already been prepared. Skipping.')
            continue

        driver = create_driver()

        try:

            """
            LOGGING IN AND ADDING RECOVERY EMAIL
            """

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

            try:
                WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable(
                        (By.XPATH, '//button[normalize-space()="Yes"]')
                    )
                ).click()
            except Exception:
                logger.critical(f'[{index+1}/{len(entries)}] Something went wrong with {entry.get("Email")}. Skipping.')
                quit_driver()
                continue

            if not entry.get('Recovery Email'):

                logger.debug(f'[{index+1}/{len(entries)}] Waiting for recovery email input.')

                recovery_email_address = create_mailbox().get('email_address')

                if not recovery_email_address:
                    logger.critical(f'[{index+1}/{len(entries)}] Failed to create mailbox. Something is wrong.')
                    quit_driver()
                    continue

                logger.debug(f'[{index+1}/{len(entries)}] Created temporary email: {recovery_email_address}')

                time.sleep(1)

                driver.get(JUNK_EMAIL)
                
                WebDriverWait(driver, 60).until(
                    EC.visibility_of_element_located(
                        (By.CSS_SELECTOR, '#EmailAddress')
                    )
                ).send_keys(recovery_email_address)
                time.sleep(0.2)
                
                WebDriverWait(driver, 5).until(
                    EC.element_to_be_clickable(
                        (By.CSS_SELECTOR, '#iNext')
                    )
                ).click()

                logger.debug(f'[{index+1}/{len(entries)}] Fetching OTP.')

                verification_code = None

                count = 0

                while True:
                    count += 1
                    time.sleep(7.5)
                    logger.debug(f'[{index+1}/{len(entries)}] Fetching emails. Attempt {count}')
                    emails = fetch_emails(recovery_email_address).get('emails')
                    if emails:
                        for email in emails:
                            if 'microsoft' in email.get('from_address'):
                                extracted = extract_microsoft_otp(email.get('content'))
                                if extracted:
                                    email_address = extracted.split('-')[0].strip('[').strip(']').strip()
                                    code = extracted.split('-')[-1].strip()

                                    email_start = email_address.split('@')[0].split('*')[0]
                                    email_end = email_address.split('@')[0].split('*')[-1]

                                    if entry.get('Email').split('@')[0].startswith(email_start) and entry.get('Email').split('@')[0].endswith(email_end):
                                        verification_code = code
                                        break

                    if verification_code:
                        break

                logger.debug(f'[{index+1}/{len(entries)}] OTP: {verification_code}')

                WebDriverWait(driver, 10).until(
                    EC.visibility_of_element_located(
                        (By.CSS_SELECTOR, '#iOttText')
                    )
                ).send_keys(verification_code)
                time.sleep(0.2)

                WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable(
                        (By.CSS_SELECTOR, '#iNext')
                    )
                ).click()
                time.sleep(0.2)

                update_entry(entry.get('id'), {'Recovery Email': recovery_email_address})

                logger.debug(f'[{index+1}/{len(entries)}] Recovery email added.')

            logger.debug(f'[{index+1}/{len(entries)}] Logged in successfully.')
            
            """
            ADDING SAFE SENDER
            """

            driver.get(JUNK_EMAIL)

            try:
                WebDriverWait(driver, 15).until(
                    EC.element_to_be_clickable(
                        (By.XPATH, '//button[normalize-space()="No, thanks"]')
                    )
                ).click()
            except Exception:
                pass

            if not entry.get('Safe Sender Added'):
                logger.debug(f'[{index+1}/{len(entries)}] Adding safe sender.')
                driver.get(JUNK_EMAIL)

                # Adding FEELD email

                WebDriverWait(driver, 60).until(
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

                # Adding TAIMI email

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
                )[-1].send_keys(TAIMI_EMAIL)
                time.sleep(0.2)

                WebDriverWait(driver, 5).until(
                    EC.element_to_be_clickable(
                        (By.XPATH, '//button[normalize-space()="OK"]')
                    )
                ).click()
                time.sleep(0.2)

                # Adding kartik.aggarwal117@gmail.com

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
                )[-1].send_keys('kartik.aggarwal117@gmail.com')
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

            """
            ADDING FORWARDING EMAIL
            """

            if not entry.get('Forwarded Email'):

                logger.debug(f'[{index+1}/{len(entries)}] Adding forwarding email.')

                driver.get(FORWARDING)

                WebDriverWait(driver, 60).until(
                    EC.element_to_be_clickable(
                        (By.XPATH, '//button[normalize-space()="Add new rule"]')
                    )
                ).click()
                time.sleep(0.2)

                # Name

                WebDriverWait(driver, 5).until(
                    EC.presence_of_element_located(
                        (By.CSS_SELECTOR, '.ms-TextField-field')
                    )
                ).send_keys('Forward all emails')
                time.sleep(0.2)

                # Condition

                WebDriverWait(driver, 5).until(
                    EC.element_to_be_clickable(
                        (By.CSS_SELECTOR, '.ms-ComboBox-CaretDown-button')
                    )
                ).click()
                time.sleep(0.2)

                WebDriverWait(driver, 5).until(
                    EC.element_to_be_clickable(
                        (By.XPATH, '//button[normalize-space()="Apply to all messages"]')
                    )
                ).click()
                time.sleep(0.2)

                # Action

                WebDriverWait(driver, 5).until(
                    EC.presence_of_all_elements_located(
                        (By.CSS_SELECTOR, '.ms-ComboBox-CaretDown-button')
                    )
                )[-1].click()
                time.sleep(0.2)

                WebDriverWait(driver, 5).until(
                    EC.element_to_be_clickable(
                        (By.XPATH, '//button[normalize-space()="Forward to"]')
                    )
                ).click()
                time.sleep(0.2)

                value_field = WebDriverWait(driver, 60).until(
                    EC.presence_of_element_located(
                        (By.CSS_SELECTOR, '.EditorClass')
                    )
                )

                value_field.click()
                time.sleep(0.2)

                pyperclip.copy(EMAIL_USER)

                value_field.send_keys(Keys.CONTROL, 'v')

                WebDriverWait(driver, 5).until(
                    EC.element_to_be_clickable(
                        (By.XPATH, '//button[normalize-space()="Save"]')
                    )
                ).click()
                time.sleep(3)

                update_entry(entry.get('id'), {'Forwarded Email': EMAIL_USER})

                logger.debug(f'[{index+1}/{len(entries)}] Forwarding email added.')

            update_entry(entry.get('id'), {'Prepared': True})

            quit_driver()
        
        except Exception as e:
            logger.error(f'[{index+1}/{len(entries)}] Something went wrong with {entry.get("Email")}. Skipping.')
            logger.error(f'\n{traceback.format_exc()}')
            quit_driver()
            continue


if __name__ == '__main__':
    main()