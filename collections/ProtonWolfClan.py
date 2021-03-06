from ast import Pass
import json
import os
from wsgiref.simple_server import WSGIServer
import pandas as pd
import requests as requests
import pathlib
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from datetime import datetime
import time

print(
    "Thank you for running the program\nIt will take a few seconds to create everything\nHope you enjoy  \n-kevcollector"
)
# set what command you want to run here
# os.getenv('PWD')
# mac=pathlib.Path().cwd() /'Desktop'
# os.chdir(mac)
# os.system('cls' if os.name == 'nt' else 'clear')
# userCollection = input("Please enter the collection you want to scan")
# os.system('cls' if os.name == 'nt' else 'clear')


def writeToExcel(worksheet, data):
    for r in dataframe_to_rows(data, index=False):
        worksheet.append(r)
    dims = {}
    for row in worksheet.rows:
        for cell in row:
            if cell.value:
                dims[cell.column_letter] = max(
                    (dims.get(cell.column_letter, 0), len(str(cell.value)))
                )
    for col, value in dims.items():
        worksheet.column_dimensions[col].width = value


current = pathlib.Path().cwd()

totalholderslist = []


def collection(author, collection_name, heading, *excelsheetname):
    collecion = "".join(excelsheetname)
    collecion = collecion.replace(".xlsx", "")

    path = pathlib.Path().cwd() / ("{}".format(heading))
    if current != pathlib.Path().cwd():
        path = pathlib.Path().cwd()
        pathlib.Path(path).mkdir(parents=True, exist_ok=True)
    else:
        pathlib.Path(path).mkdir(parents=True, exist_ok=True)

    holders = " "

    holder_list = []

    def normalServic(author, holders, *excelsheetname):
        global totalholderslist
        os.chdir(path)

        wb1 = Workbook()
        ws1 = wb1.active
        ws1.title = "holders"

        pages = 0
        holders = (
            "https://proton.api.atomicassets.io/atomicassets/v1/accounts?collection_name={}"
            "&page=1&limit=100&order=desc".format(collection_name)
        )
        holders = requests.get(holders).text
        holders_ = json.loads(holders)
        amount = 1
        rowz = 1

        while len(holders_["data"]) != 0:
            pages = pages + 1
            amount = amount + 1
            for data_info in holders_["data"]:
                if data_info["account"] != author:
                    checker = data_info["account"]
                    print("getting {}'s data".format(checker))
                    people = "https://proton.api.atomicassets.io/atomicmarket/v1/assets?collection_name={}&owner={}&page=1&limit=100&order=desc&sort=asset_id".format(
                        collection_name, checker
                    )
                    test = requests.get(people)
                    next = test.headers["X-RateLimit-Reset"]
                    resset = test.headers["X-RateLimit-Remaining"]
                    resset = int(resset)
                    next = int(next)
                    wait = next - time.time()
                    if resset < 3:
                        time.sleep(wait)
                    people_ = json.loads((test.text))
                    count = 0
                    wolfPoints = 0
                    rowz = rowz + 1
                    pages = 1
                    assitID2 = " "

                    while len(people_["data"]) != 0:
                        pages = pages + 1
                        if pages >= 3:
                            count = pages * 100

                        for data_info in people_["data"]:
                            nft_name = data_info["data"]["name"]
                            assitID1 = data_info["asset_id"]
                            if assitID1 != assitID2:
                                assitID2 = assitID1
                                done = nft_name

                                # if collection_name == "524211545444":
                                # insert lookup table here

                                if collection_name == "3drfwzczslri":
                                    count = count + 1

                                    if "COMMON" in nft_name:
                                        ws1.cell(
                                            row=rowz, column=3
                                        ).value = wolfPoints = (wolfPoints + 2)

                                    if "RARE" in nft_name:
                                        ws1.cell(
                                            row=rowz, column=3
                                        ).value = wolfPoints = (wolfPoints + 5)

                                    if "EPIC" in nft_name:
                                        ws1.cell(
                                            row=rowz, column=3
                                        ).value = wolfPoints = (wolfPoints + 7)

                                    if "HEROIC" in nft_name:
                                        ws1.cell(
                                            row=rowz, column=3
                                        ).value = wolfPoints = (wolfPoints + 9)

                                    if "ULTRA RARE" in nft_name:
                                        ws1.cell(
                                            row=rowz, column=3
                                        ).value = wolfPoints = (wolfPoints + 10)

                                    if "ULTRA EPIC" in nft_name:
                                        ws1.cell(
                                            row=rowz, column=3
                                        ).value = wolfPoints = (wolfPoints + 14)

                                    if "LEGENDARY" in nft_name:
                                        ws1.cell(
                                            row=rowz, column=3
                                        ).value = wolfPoints = (wolfPoints + 15)

                                    if "FUSION" in nft_name:
                                        if "ULTRA RARE" in nft_name:
                                            ws1.cell(
                                                row=rowz, column=3
                                            ).value = wolfPoints = (wolfPoints + 10)

                                        if "ULTRA EPIC" in nft_name:
                                            ws1.cell(
                                                row=rowz, column=3
                                            ).value = wolfPoints = (wolfPoints + 14)
                                    else:
                                        ws1.cell(row=rowz, column=3).value = wolfPoints

                                    ws1.cell(row=rowz, column=3 + count).value = done

                                if count > 0:
                                    ws1.cell(
                                        row=1, column=3 + count
                                    ).value = "NFT " + str(count)
                                    ws1.cell(row=1, column=1).value = "Account"
                                    ws1.cell(row=1, column=2).value = "Amount"
                                    ws1.cell(row=1, column=3).value = "Points"
                                    ws1.cell(row=rowz, column=1).value = checker
                                    ws1.cell(row=rowz, column=2).value = count

                        people = "https://proton.api.atomicassets.io/atomicmarket/v1/assets?collection_name={}&owner={}&page={}&limit=100&order=desc&sort=asset_id".format(
                            collection_name, checker, pages
                        )
                        test = requests.get(people)
                        next = test.headers["X-RateLimit-Reset"]
                        resset = test.headers["X-RateLimit-Remaining"]
                        resset = int(resset)
                        next = int(next)
                        wait = next - time.time()
                        if resset < 3:
                            time.sleep(wait)
                        people_ = json.loads((test.text))

                    holders_amount = ws1.cell(row=rowz, column=2).value

                    totalPoints = ws1.cell(row=rowz, column=3).value

                    totalholderslist.append([checker, holders_amount, totalPoints])

            holders = (
                "https://proton.api.atomicassets.io/atomicassets/v1/accounts?collection_name={}"
                "&page={}&limit=100&order=desc".format(collection_name, amount)
            )
            holders = requests.get(holders).text
            holders_ = json.loads(holders)

        count = 0
        dims = {}
        for row in ws1.rows:
            for cell in row:
                if cell.value:
                    dims[cell.column_letter] = max(
                        (dims.get(cell.column_letter, 0), len(str(cell.value)))
                    )
        for col, value in dims.items():
            ws1.column_dimensions[col].width = value + 5

        excelsave = "".join(excelsheetname)
        wb1.save(excelsave)
        print("Creating the excel file")

        wb1.close()

        os.chdir(path.parent.absolute())

    normalServic(author, holders, *excelsheetname)


author = "redbrush"
universe = "Proton Wolf Clan"
heading = "{} Collection".format(universe)
# collection_name = "524211545444"
# collection1 = "PWC Gen I"
# excelsheetname1 = "{}.xlsx".format(collection1)
# time.sleep(6)
# collection(author, collection_name, heading, excelsheetname1)
collection_name = "3drfwzczslri"
collection2 = "PWC Gen II"
excelsheetname1 = "{}.xlsx".format(collection2)
time.sleep(6)
collection(author, collection_name, heading, excelsheetname1)
wb2 = Workbook()
ws1 = wb2.active
ws1 = wb2.create_sheet("holders")
holders_df = pd.DataFrame(
    data=totalholderslist, columns=["account", "amount held", "points"]
)

print(totalholderslist)

holders_df = holders_df.sort_values(by=["points"], ascending=False)


writeToExcel(ws1, holders_df)

wb2.save("testfile.xlsx")
