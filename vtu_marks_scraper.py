import time
import os
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (UnexpectedAlertPresentException,
                                        NoAlertPresentException,
                                        NoSuchElementException)
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from Analyzer import analyze_results



def generate_usn_list(base="1CR24BA", start=1, end=10):
    """Generate list of USNs with given base and numeric range"""
    return [f"{base}{str(i).zfill(3)}" for i in range(start, end + 1)]


def get_driver(headless=True):
    """Initialize and return a Chrome WebDriver with proper options"""
    options = Options()
    if headless:
        options.add_argument("--headless")
    #options.add_argument("--disable-gpu")
    #options.add_argument("--no-sandbox")
    #options.add_argument("--disable-dev-shm-usage")
    #options.add_argument("--log-level=3")
    options.add_experimental_option('excludeSwitches', ['enable-logging'])

    try:
        driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
        driver.set_page_load_timeout(30)
        return driver
    except Exception as e:
        driver = webdriver.Chrome(options=options)
        return driver
    except:
        print(f"Failed to initialize WebDriver: {e}")
        raise



def fetch_vtu_result_with_retry(driver, usn, captcha_handler, max_retries=50, base_url=None):
    """Fetch VTU result with retry mechanism"""
    attempt = 1
    while attempt <= max_retries:
        try:
            print(f"[Attempt {attempt}/{max_retries}] Processing USN: {usn}")

            # Load the result page
            driver.get(base_url)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.NAME, "lns")))

            # Enter USN
            usn_input = driver.find_element(By.NAME, "lns")
            usn_input.clear()
            usn_input.send_keys(usn)

            # CAPTCHA solving
            captcha_valid = False
            captcha_retries = 0
            max_captcha_retries = 3

            while not captcha_valid and captcha_retries < max_captcha_retries:
                captcha_element = driver.find_element(By.XPATH, '//*[@alt="CAPTCHA code"]')
                captcha_png = captcha_element.screenshot_as_png
                captcha_text = captcha_handler.get_captcha_from_image(captcha_png).strip()
                print(f"Solved CAPTCHA: {captcha_text} (Length: {len(captcha_text)})")

                if len(captcha_text) == 6:
                    captcha_valid = True
                else:
                    captcha_retries += 1
                    print(f"CAPTCHA too short, retrying... ({captcha_retries}/{max_captcha_retries})")
                    #driver.find_element(By.XPATH, '//a[contains(text(), "Refresh")]').click()

            if not captcha_valid:
                print("Failed to get valid CAPTCHA after retries")
                attempt += 1
                continue

            # Fill CAPTCHA and submit
            captcha_input = driver.find_element(By.NAME, 'captchacode')
            captcha_input.clear()
            captcha_input.send_keys(captcha_text)
            driver.find_element(By.ID, "submit").click()

            WebDriverWait(driver, 10).until(
                EC.any_of(
                    EC.presence_of_element_located((By.XPATH, '//div[@class="panel-body"]/div[@class="row"][1]')),
                    EC.alert_is_present()
                )
            )

            try:
                alert = driver.switch_to.alert
                alert_text = alert.text.strip()
                alert.accept()

                if "University Seat Number is not available or Invalid" in alert_text:
                    print("Invalid USN. Skipping further attempts.")
                    return None

                elif "Invalid captcha code !!!" in alert_text:
                    print(f"[CAPTCHA error Attempt {attempt}] Failed : Retrying")
                    attempt += 1
                    continue

            except NoAlertPresentException:
                pass

            # Extract result content
            time.sleep(2)
            result_container = driver.find_element(By.XPATH, '//div[@class="panel-body"]/div[@class="row"][1]')
            html_content = result_container.get_attribute('outerHTML')
            print("Successfully fetched result")
            return html_content

        except Exception as e:
            print(f"Error: [Attempt {attempt}] Failed: {str(e)}")
            attempt += 1
            if attempt <= max_retries:
                print("Retrying.....")

    print(f"All {max_retries} attempts failed for USN: {usn}")
    return None






