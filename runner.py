import os
import sys
import pathlib
from pyWinActivate import win_activate, win_wait_active
import keyboard
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import modules.ApiClass as Api
import modules.SalesClass as Sale
import modules.transfersclass as t
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select

a = "b"

author = "kevcollector"
collection_name = "513222111131"

authors = Api.ApiAuthor(author, collection_name)
sales = Sale.Sales(authors.authors_)
sales.showSale
print(sales)
print(sales.lists)
sales.showSales
""""
if getattr(sys, 'frozen', False):
    firefoxpath = os.path.join(sys._MEIPASS, "geckodriver.exe")
    driver = webdriver.Firefox()
else:
    driver = webdriver.Firefox()
"""
userInput = input("Please enter your wallet pin/ windows")
walletname = "kevcollector"


def checkwin(userInput):
    f = win_activate(window_title="Windows Security", partial_match=True)
    time.sleep(0.2)
    keyboard.write(userInput)


def trys(wayx, patj):
    if wayx == "xpath":
        way = (By.XPATH, patj)
    if wayx == "id":
        way = (By.ID, patj)
    if wayx == "name":
        way = (By.NAME, patj)

    element = WebDriverWait(driver, 10).until(EC.presence_of_element_located(way))
    element.click()


def tryAdd(wayx, patj, number: str):
    if wayx == "xpath":
        way = (By.XPATH, patj)
    if wayx == "id":
        way = (By.ID, patj)
    if wayx == "name":
        way = (By.NAME, patj)
        element = WebDriverWait(driver, 10).until(EC.presence_of_element_located(way))
        element.clear()
        element.send_keys(number + Keys.ENTER)


def login(walletname, driver):

    login = driver.find_element(By.XPATH, '//*[@id="actor"]')
    login.send_keys(walletname + Keys.ENTER)
    trys(
        "xpath",
        "/html/body/div[1]/div/div/div/main/div/div/div/div[2]/div/div/div[2]/button",
    )
    time.sleep(1.5)
    driver.get("https://www.protonscan.io/wallet/transfer")
    time.sleep(1)
    window_before = driver.window_handles[0]
    trys("xpath", "/html/body/div[1]/div[1]/div[1]/div/div/div/div[6]/div/div/span")
    trys(
        "xpath",
        "/html/body/div[1]/div[1]/div[3]/div/div/div[2]/div/div[2]/div/div[1]/div",
    )
    trys("xpath", "/html/body/div[3]/div[1]/div[1]/div[2]/ul/li[2]")
    time.sleep(1)
    return window_before


driver = webdriver.Firefox()
driver.get("https://webauth.com/login")

# click the login on protonscan
window_before = login(walletname, driver)

for x in sales.lists:
    tryAdd(
        "xpath",
        "/html/body/div[1]/div[2]/div/div[2]/div[3]/div[2]/div/div/div/div[1]/input",
        sales.buyer,
    )
    trys(
        "xpath",
        "/html/body/div[1]/div[2]/div/div[2]/div[3]/div[2]/div/div/div/div[2]/div[2]/a[2]",
    )
    if sales.coin == "XUSDC":
        amount = sales.moneyUSDC * 0.1
        tryAdd(
            "xpath",
            "/html/body/div[1]/div[2]/div/div[2]/div[3]/div[2]/div/div/div/div[3]/div[1]/input",
            "XUSDC",
        )
        tryAdd(
            "xpath",
            "/html/body/div[1]/div[2]/div/div[2]/div[3]/div[2]/div/div/div/div[3]/div[2]/input",
            "xtokens",
        )
        tryAdd(
            "xpath",
            "/html/body/div[1]/div[2]/div/div[2]/div[3]/div[2]/div/div/div/div[2]/div[1]/input",
            amount,
        )

    if sales.coin == "XPR":
        amount = sales.moneyXPR * 0.1
        tryAdd(
            "xpath",
            "/html/body/div[1]/div[2]/div/div[2]/div[3]/div[2]/div/div/div/div[3]/div[1]/input",
            "XPR",
        )
        tryAdd(
            "xpath",
            "/html/body/div[1]/div[2]/div/div[2]/div[3]/div[2]/div/div/div/div[3]/div[2]/input",
            "eosio.token",
        )
        tryAdd(
            "xpath",
            "/html/body/div[1]/div[2]/div/div[2]/div[3]/div[2]/div/div/div/div[2]/div[1]/input",
            str(amount),
        )
    if sales.moneyLOAN > 1:
        moneyLOAN = sales.moneyLOAN * 0.1
        tryAdd(
            "xpath",
            "/html/body/div[1]/div[2]/div/div[2]/div[3]/div[2]/div/div/div/div[3]/div[1]/input",
            "LOAN",
        )
        tryAdd(
            "xpath",
            "/html/body/div[1]/div[2]/div/div[2]/div[3]/div[2]/div/div/div/div[3]/div[2]/input",
            "loan.token",
        )
        tryAdd(
            "xpath",
            "/html/body/div[1]/div[2]/div/div[2]/div[3]/div[2]/div/div/div/div[2]/div[1]/input",
            str(moneyLOAN),
        )

    time.sleep(0.1)
    moneyA = sales.moneyUSDC * 1
    if moneyA < 1:
        pass
    else:
        tryAdd(
            "xpath",
            "html/body/div[1]/div[2]/div/div[2]/div[3]/div[2]/div/div/div/div[2]/div[1]/input",
            str(sales.moneyUSDC),
        )
    time.sleep(1)
    trys(
        "xpath", "/html/body/div[1]/div[2]/div/div[2]/div[3]/div[2]/div/div/div/button"
    )
    window_after = driver.window_handles[1]
    time.sleep(1)
    driver.switch_to.window(window_after)
    time.sleep(1)
    trys("xpath", "/html/body/div/div/div/div[2]/main/div/div/div[2]/button[2]")
    time.sleep(1)
    checkwin(userInput)
    time.sleep(0.5)

    driver.switch_to.window(window_before)
    #
