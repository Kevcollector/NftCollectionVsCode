import os
import sys
import pathlib
import pywinauto
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
""""
if getattr(sys, 'frozen', False):
    firefoxpath = os.path.join(sys._MEIPASS, "geckodriver.exe")
    driver = webdriver.Firefox()
else:
    driver = webdriver.Firefox()
"""


def security(password):
    while True:

        try:
            app = Application(backend="uia").connect(title="Windows Security")
            app.TypeKeys(password)
        except (pywinauto.findwindows.WindowNotFoundError, pywinauto.timings.TimeoutError):
            print("DID WORK")


def trys(wayx, patj):
    if wayx == "xpath":
        way = (By.XPATH, patj)
    if wayx == "id":
        way = (By.ID, patj)
    if wayx == "name":
        way = (By.NAME, patj)

    element = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located(way)
    )
    element.click()


def tryAdd(wayx, patj, number: str):
    if wayx == "xpath":
        way = (By.XPATH, patj)
    if wayx == "id":
        way = (By.ID, patj)
    if wayx == "name":
        way = (By.NAME, patj)
    try:
        element = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located(way)
        )
        element.click()
        element.clear()
        element.send_keys(number+Keys.ENTER)
    except:
        print("something bad happened")


current = pathlib.Path().cwd()
print(current)
driver = webdriver.Firefox(current)
driver.get("https://webauth.com/login")

login = driver.find_element(By.XPATH, '//*[@id="actor"]')
login.send_keys("kevcollector" + Keys.ENTER)
trys("xpath",
     "/html/body/div[1]/div/div/div/main/div/div/div/div[2]/div/div/div[2]/button")
time.sleep(1)
driver.get("https://www.protonscan.io/wallet/transfer")
window_before = driver.window_handles[0]

# click the login on protonscan
trys(
    "xpath", "/html/body/div[1]/div[1]/div[1]/div/div/div/div[6]/div/div/span")
trys("xpath",
     "/html/body/div[1]/div[1]/div[3]/div/div/div[2]/div/div[2]/div/div[1]/div")
trys("xpath", "/html/body/div[3]/div[1]/div[1]/div[2]/ul/li[2]")
time.sleep(.4)
tryAdd("xpath",
       "/html/body/div[1]/div[2]/div/div[2]/div[3]/div[2]/div/div/div/div[1]/input", "anderson22")
tryAdd("xpath",
       "html/body/div[1]/div[2]/div/div[2]/div[3]/div[2]/div/div/div/div[2]/div[1]/input", ".001")
time.sleep(1.5)

time.sleep(1.5)
driver.find_element(
    By.XPATH, "/html/body/div[1]/div[2]/div/div[2]/div[3]/div[2]/div/div/div/button").click()
time.sleep(1)
window_after = driver.window_handles[1]
driver.switch_to.window(window_after)
time.sleep(1)
auth = driver.find_element(
    By.XPATH, "/html/body/div/div/div/div[2]/main/div/div/div[2]/button[2]")
auth.click()
security("4517")
time.sleep(1)
for x in range(1, 10):
    driver.switch_to.window(window_before)

    time.sleep(2)
    for window_handle in driver.window_handles:
        if window_handle != window_before:
            driver.switch_to.window(window_handle)
            break
    time.sleep(1)
    auth = driver.find_element(
        By.XPATH, "/html/body/div/div/div/div[2]/main/div/div/div[2]/button[2]")
    auth.click()
    os.system()
    time.sleep(1)


#
