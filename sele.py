from selenium import webdriver # type: ignore
from selenium.webdriver.common.keys import Keys # type: ignore
import time

# Path to your WebDriver
driver_path = '/chromedriver-win64/chromedriver-win64'
driver = webdriver.Chrome(driver_path)

# Open WhatsApp Web
driver.get('https://web.whatsapp.com')
input("Press Enter after scanning QR code and WhatsApp Web is loaded completely")

def send_message(phone_number, message):
    driver.get(f'https://web.whatsapp.com/send?phone={phone_number}&text={message}')
    time.sleep(10)  # Wait for the page to load
    send_button = driver.find_element_by_xpath('//button[@data-icon="send"]')
    send_button.click()

# Example usage
phone_numbers = ['+6285847624457']  # Add more phone numbers as needed
message = 'Hello, this is a test message from Selenium!'

for number in phone_numbers:
    send_message(number, message)
    time.sleep(5)  # Wait between messages to avoid detection

driver.quit()
