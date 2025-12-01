import time
import argparse
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys

def send_email(email, password, message):
    # Use your downloaded ChromeDriver
    driver = webdriver.Chrome()  # <-- DO NOT ADD ANYTHING HERE
    driver.maximize_window()
    driver.get("https://outlook.live.com/")
    time.sleep(3)

    # Sign in
    driver.find_element(By.LINK_TEXT, "Sign in").click()
    time.sleep(3)

    # Enter email
    driver.find_element(By.NAME, "loginfmt").send_keys(email + Keys.ENTER)
    time.sleep(3)

    # Enter password
    driver.find_element(By.NAME, "passwd").send_keys(password + Keys.ENTER)
    time.sleep(3)

    # "Stay signed in?" → No
    try:
        driver.find_element(By.ID, "idBtn_Back").click()
        time.sleep(3)
    except:
        pass

    # New mail
    driver.find_element(By.XPATH, "//button[@aria-label='New mail']").click()
    time.sleep(3)

    # To
    driver.find_element(By.XPATH, "//input[@aria-label='To']").send_keys("scittest@auditram.com")
    time.sleep(1)

    # Subject
    driver.find_element(By.XPATH, "//input[@placeholder='Add a subject']").send_keys("AuditRAM Assignment Email")
    time.sleep(1)

    # Body
    driver.find_element(By.XPATH, "//div[@role='textbox']").send_keys(message)
    time.sleep(1)

    # Send
    driver.find_element(By.XPATH, "//button[@aria-label='Send']").click()
    time.sleep(3)

    print("Email sent successfully!")
    driver.quit()


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--email", required=True)
    parser.add_argument("--password", required=True)
    parser.add_argument("--message", required=True)
    args = parser.parse_args()

    send_email(args.email, args.password, args.message)
